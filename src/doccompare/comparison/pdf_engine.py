"""PDF-native document comparison engine.

Compares two PDF files by extracting text (via pdfplumber), diffing paragraphs
with the same weighted-LCS + diff_match_patch approach as the OOXML engine,
and producing output with tracked-changes styling.

Architecture:
  1. Extract paragraphs from both PDFs (pdfplumber for text + char info)
  2. Clean up: filter headers/footers, merge cross-page paragraphs
  3. Strip leading numbering before matching (avoids renumbering noise)
  4. Match paragraphs via weighted LCS (same algorithm as ooxml_engine)
  5. Character-level diff on matched pairs (diff_match_patch)
  6. Clone the new PDF as base, mark additions (blue underline),
     insert deletions (red strikethrough) — preserves original formatting
  7. Append summary/legend page
"""

import html as html_mod
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import diff_match_patch as dmp_module
from loguru import logger
from rapidfuzz import fuzz

# -- Regex patterns -----------------------------------------------------------

# Leading numbering: "1.12 ", "12.3.4 ", "(a) ", "(iv) ", "a) " etc.
_RE_LEADING_NUM = re.compile(
    r"^(?:"
    r"\d+(?:\.\d+)*\.?\s+"           # 1.12 or 1.12. or 12.3.4
    r"|\(\w+\)\s+"                    # (a) or (iv) or (1)
    r"|\w+\)\s+"                      # a) or iv)
    r")"
)

# Headers/footers: document IDs, page numbers, "MATTER X | Y"
_RE_HEADER_FOOTER = re.compile(
    r"^\d{5,}v?\d*\s+MATTER\s+\d+"   # "22846400v1 MATTER 3 | 21"
    r"|^Page\s+\d+\s+of\s+\d+"       # "Page 1 of 42"
    r"|^\d+\s*$"                       # bare page number
    r"|^\d+\s*\|\s*\d+"              # "21 | 42" style page numbers
    r"|^[-\u2013\u2014]\s*\d+\s*[-\u2013\u2014]"   # "- 21 -" style
    r"|MATTER\s+\d+\s*\|\s*\d+"      # "MATTER 3 | 21" anywhere
, re.IGNORECASE)


# -- Diff helpers -------------------------------------------------------------

def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return fuzz.ratio(a, b) / 100.0


def _strip_numbering(text: str) -> str:
    """Strip leading numbering from text for comparison purposes."""
    return _RE_LEADING_NUM.sub("", text)


def _match_paragraphs(old_paras: list[str], new_paras: list[str]) -> list[tuple[int, int]]:
    """Weighted LCS matching of paragraph texts."""
    n, m = len(old_paras), len(new_paras)

    old_stripped = [_strip_numbering(t) for t in old_paras]
    new_stripped = [_strip_numbering(t) for t in new_paras]

    sim_cache: dict = {}
    for i in range(n):
        for j in range(m):
            s = _similarity(old_stripped[i], new_stripped[j])
            if s > 0.4:
                sim_cache[(i, j)] = s

    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            s = sim_cache.get((i - 1, j - 1))
            if s is not None:
                dp[i][j] = max(dp[i - 1][j - 1] + s, dp[i - 1][j], dp[i][j - 1])
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        s = sim_cache.get((i - 1, j - 1))
        if s is not None and abs(dp[i][j] - (dp[i - 1][j - 1] + s)) < 1e-9:
            matches.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] > dp[i][j - 1] + 1e-9:
            i -= 1
        else:
            j -= 1
    matches.reverse()
    return matches


def _diff_paragraph(old_text: str, new_text: str) -> list[tuple[str, str]]:
    """Character-level diff returning [(type, text), ...].

    Diffs on the numbering-stripped body, then prepends numbering changes
    as a single unit (avoids character-level noise on "1.11" -> "1.12").
    """
    old_num_m = _RE_LEADING_NUM.match(old_text)
    new_num_m = _RE_LEADING_NUM.match(new_text)

    old_num = old_num_m.group() if old_num_m else ""
    new_num = new_num_m.group() if new_num_m else ""
    old_body = old_text[len(old_num):]
    new_body = new_text[len(new_num):]

    result = []

    # Handle numbering change as a single unit
    if old_num == new_num:
        if old_num:
            result.append(("equal", old_num))
    else:
        if old_num:
            result.append(("deleted", old_num))
        if new_num:
            result.append(("added", new_num))

    # Diff the body text
    dmp = dmp_module.diff_match_patch()
    diffs = dmp.diff_main(old_body, new_body)
    dmp.diff_cleanupSemantic(diffs)

    for op, text in diffs:
        if op == 0:
            result.append(("equal", text))
        elif op == 1:
            result.append(("added", text))
        elif op == -1:
            result.append(("deleted", text))
    return result


# -- Paragraph extraction & cleanup ------------------------------------------

def _extract_paragraphs(pdf_path: Path) -> list[dict]:
    """Extract paragraphs from a PDF with text and formatting info."""
    import pdfplumber

    raw_paras = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        num_pages = len(pdf.pages)
        page_width = pdf.pages[0].width if pdf.pages else 595

        # Determine left margin from the most common x0 across all pages
        all_x0 = []
        for page in pdf.pages:
            for line in (page.extract_text_lines() or []):
                all_x0.append(round(line.get("x0", 0), 0))

        if all_x0:
            x0_counts = Counter(all_x0)
            common_x0 = sorted(x0_counts.items(), key=lambda x: -x[1])
            left_margin = min(x for x, _ in common_x0[:3])
        else:
            left_margin = 71.0

        for page_num, page in enumerate(pdf.pages):
            lines = page.extract_text_lines() or []
            if not lines:
                continue

            paragraphs = _group_lines(lines, page_width, left_margin)

            for para in paragraphs:
                para_text = para["text"].strip()
                if not para_text:
                    continue
                para["page_num"] = page_num
                para["element_type"] = "paragraph"
                raw_paras.append(para)

    # Step 1: Detect and filter headers/footers
    paras = _filter_headers_footers(raw_paras, num_pages)

    # Step 2: Merge paragraphs broken across page boundaries
    paras = _merge_cross_page(paras)

    logger.debug(f"After cleanup: {len(paras)} paragraphs (from {len(raw_paras)} raw)")
    return paras


def _detect_font_style(chars: list) -> tuple[bool, bool]:
    """Detect bold/italic from character font names."""
    if not chars:
        return False, False

    bold_count = 0
    italic_count = 0
    total = 0

    for c in chars:
        fn = c.get("fontname", "").lower()
        if not fn or c.get("text", "").strip() == "":
            continue
        total += 1
        if "bold" in fn:
            bold_count += 1
        if "italic" in fn or "oblique" in fn:
            italic_count += 1

    if total == 0:
        return False, False

    return (bold_count / total > 0.5), (italic_count / total > 0.5)


def _group_lines(lines: list, page_width: float, left_margin: float) -> list[dict]:
    """Group text lines into paragraphs based on vertical spacing."""
    if not lines:
        return []

    paragraphs = []
    current_lines = []
    prev_bottom = None

    def _flush():
        if not current_lines:
            return
        text = " ".join(l.get("text", "") for l in current_lines)

        all_sizes = []
        all_chars = []
        for l in current_lines:
            chars = l.get("chars", [])
            all_chars.extend(chars)
            for c in chars:
                s = c.get("size")
                if s:
                    all_sizes.append(s)
        avg_size = sum(all_sizes) / len(all_sizes) if all_sizes else 11.0

        is_bold, is_italic = _detect_font_style(all_chars)

        first_x0 = current_lines[0].get("x0", left_margin)
        indent_pt = max(0, first_x0 - left_margin)

        first_line = current_lines[0]
        x0 = first_line.get("x0", 0)
        x1 = first_line.get("x1", page_width)
        text_center = (x0 + x1) / 2
        page_center = page_width / 2
        is_centered = (
            abs(text_center - page_center) < 30
            and (x0 - left_margin) > 30
            and len(current_lines) <= 3
        )

        paragraphs.append({
            "text": text,
            "font_size": round(avg_size, 1),
            "is_bold": is_bold,
            "is_italic": is_italic,
            "indent_pt": round(indent_pt, 0),
            "is_centered": is_centered,
            "level": 0,
        })

    for line in lines:
        top = line.get("top", 0)
        bottom = line.get("bottom", top + 12)
        height = bottom - top

        if prev_bottom is not None:
            gap = top - prev_bottom
            if gap > height * 0.8:
                _flush()
                current_lines = []

        if line.get("text", ""):
            current_lines.append(line)

        prev_bottom = bottom

    _flush()
    return paragraphs


def _filter_headers_footers(paras: list[dict], num_pages: int) -> list[dict]:
    """Remove header/footer lines that repeat across pages."""
    if num_pages < 2:
        return paras

    text_page_sets: dict[str, set[int]] = {}
    for p in paras:
        t = p["text"].strip()
        if len(t) > 80:
            continue
        normalized = re.sub(r"\d+", "#", t)
        if normalized not in text_page_sets:
            text_page_sets[normalized] = set()
        text_page_sets[normalized].add(p["page_num"])

    threshold = max(2, num_pages * 0.4)
    repeated_patterns = {
        pat for pat, pages in text_page_sets.items()
        if len(pages) >= threshold
    }

    filtered = []
    removed = 0
    for p in paras:
        t = p["text"].strip()
        if _RE_HEADER_FOOTER.search(t):
            removed += 1
            continue
        if len(t) <= 80:
            normalized = re.sub(r"\d+", "#", t)
            if normalized in repeated_patterns:
                removed += 1
                continue
        filtered.append(p)

    if removed:
        logger.debug(f"Filtered {removed} header/footer paragraphs")
    return filtered


def _merge_cross_page(paras: list[dict]) -> list[dict]:
    """Merge paragraphs that were split at page boundaries."""
    if len(paras) < 2:
        return paras

    merged = [paras[0]]
    for i in range(1, len(paras)):
        prev = merged[-1]
        curr = paras[i]

        if prev["page_num"] != curr["page_num"] and curr["page_num"] == prev["page_num"] + 1:
            prev_text = prev["text"].rstrip()
            curr_text = curr["text"].lstrip()

            if (prev_text and curr_text
                    and prev_text[-1] not in ".!?:;\"'"
                    and curr_text[0].islower()):
                prev["text"] = prev_text + " " + curr_text
                continue

        merged.append(curr)

    if len(merged) < len(paras):
        logger.debug(f"Merged {len(paras) - len(merged)} cross-page paragraph breaks")
    return merged


# -- Summary computation ------------------------------------------------------

def _compute_summary(old_paras: list[dict], new_paras: list[dict]) -> dict:
    """Compute word-level summary from paragraph diffs."""
    old_texts = [p["text"] for p in old_paras]
    new_texts = [p["text"] for p in new_paras]

    matches = _match_paragraphs(old_texts, new_texts)
    matched_old = {i for i, _ in matches}
    matched_new = {j for _, j in matches}

    added_words = 0
    deleted_words = 0
    unchanged_words = 0

    for oi, nj in matches:
        diff_segments = _diff_paragraph(old_paras[oi]["text"], new_paras[nj]["text"])
        for seg_type, seg_text in diff_segments:
            wc = len(seg_text.split())
            if seg_type == "equal":
                unchanged_words += wc
            elif seg_type == "added":
                added_words += wc
            elif seg_type == "deleted":
                deleted_words += wc

    for i in range(len(old_paras)):
        if i not in matched_old:
            deleted_words += len(old_paras[i]["text"].split())

    for j in range(len(new_paras)):
        if j not in matched_new:
            added_words += len(new_paras[j]["text"].split())

    return {
        "added_words": added_words,
        "deleted_words": deleted_words,
        "unchanged_words": unchanged_words,
    }


# -- PDF comparison via DOCX conversion ----------------------------------------
#
# Strategy: convert both PDFs to DOCX (via pdf2docx), then use the same
# OOXML track-changes engine that works perfectly for .docx files.
# This gives identical visual output: inline blue additions, red deletions,
# proper text reflow, and Word's native rendering.
#
# Pipeline: PDF → DOCX (pdf2docx) → cleanup section breaks → OOXML compare
#           → Word headless PDF export → merge summary page


def _pdf_to_docx(pdf_path: Path) -> Path:
    """Convert a PDF to DOCX using pdf2docx, then clean up artifacts."""
    import tempfile
    from pdf2docx import Converter

    # Create temp DOCX in a safe location
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False,
                                      prefix="doccompare_")
    tmp.close()
    docx_path = Path(tmp.name)

    logger.info(f"Converting {pdf_path.name} to DOCX via pdf2docx")
    cv = Converter(str(pdf_path))
    cv.convert(str(docx_path))
    cv.close()

    # Comprehensive cleanup of pdf2docx artifacts
    _cleanup_pdf2docx_artifacts(docx_path)

    return docx_path


# Patterns that match header/footer text embedded by pdf2docx
_RE_HF_PATTERNS = [
    re.compile(r"^\d{5,}v?\d*\s+MATTER\s+\d+", re.IGNORECASE),    # "22846400v1 MATTER 3 | 21"
    re.compile(r"^Page\s+\d+\s+of\s+\d+", re.IGNORECASE),          # "Page 1 of 42"
    re.compile(r"^\d+\s*\|\s*\d+\s*$"),                              # "21 | 42"
    re.compile(r"^[-\u2013\u2014]\s*\d+\s*[-\u2013\u2014]\s*$"),    # "- 21 -"
    re.compile(r"^MATTER\s+\d+\s*\|\s*\d+", re.IGNORECASE),         # "MATTER 3 | 21"
    re.compile(r"^\d{1,4}\s*$"),                                      # bare page number
]


def _is_header_footer_text(text: str) -> bool:
    """Check if text looks like a PDF header or footer."""
    t = text.strip()
    if not t or len(t) > 100:
        return False
    return any(pat.search(t) for pat in _RE_HF_PATTERNS)


def _cleanup_pdf2docx_artifacts(docx_path: Path):
    """Comprehensive cleanup of pdf2docx conversion artifacts.

    Fixes:
      1. Per-page w:sectPr (section breaks) → removed (keeps body-level only)
      2. Explicit page breaks (w:br type="page") → removed
      3. w:lastRenderedPageBreak → removed (rendering hint that forces breaks)
      4. w:pageBreakBefore in paragraph properties → removed
      5. Header/footer text embedded as body paragraphs → removed
      6. Empty paragraphs left behind → removed
      7. Excessive w:spacing (before/after > 400 twips) → capped
    """
    import zipfile
    import os
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(str(docx_path)) as z:
        doc_xml = z.read("word/document.xml")

    root = etree.fromstring(doc_xml)
    body = root.find(f"{{{W}}}body")
    if body is None:
        return

    stats = {"sectPr": 0, "pageBr": 0, "lastRendered": 0,
             "breakBefore": 0, "headerFooter": 0, "emptyPara": 0,
             "spacingCapped": 0}

    # --- 1. Remove paragraph-level section breaks ---
    for ppr in body.findall(f".//{{{W}}}pPr"):
        sect = ppr.find(f"{{{W}}}sectPr")
        if sect is not None:
            ppr.remove(sect)
            stats["sectPr"] += 1

    # --- 2. Remove explicit page breaks (w:br type="page") ---
    for br in body.findall(f".//{{{W}}}br"):
        br_type = br.get(f"{{{W}}}type", "")
        if br_type == "page":
            br.getparent().remove(br)
            stats["pageBr"] += 1

    # --- 3. Remove w:lastRenderedPageBreak ---
    for lrpb in body.findall(f".//{{{W}}}lastRenderedPageBreak"):
        lrpb.getparent().remove(lrpb)
        stats["lastRendered"] += 1

    # --- 4. Remove w:pageBreakBefore ---
    for pbb in body.findall(f".//{{{W}}}pageBreakBefore"):
        pbb.getparent().remove(pbb)
        stats["breakBefore"] += 1

    # --- 5. Remove header/footer paragraphs ---
    for p in list(body):
        if p.tag != f"{{{W}}}p":
            continue
        # Collect all text in the paragraph
        texts = [t.text or "" for t in p.findall(f".//{{{W}}}t")]
        full_text = "".join(texts).strip()
        if full_text and _is_header_footer_text(full_text):
            body.remove(p)
            stats["headerFooter"] += 1

    # --- 6. Remove empty paragraphs (no text, no runs, or empty pPr only) ---
    for p in list(body):
        if p.tag != f"{{{W}}}p":
            continue
        has_text = any((t.text or "").strip()
                       for t in p.findall(f".//{{{W}}}t"))
        if has_text:
            continue
        # Keep paragraphs that have drawing/image content
        if p.findall(f".//{{{W}}}drawing") or p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            continue
        runs = p.findall(f"{{{W}}}r")
        # Check if runs have any non-whitespace content
        run_has_content = False
        for r in runs:
            for child in r:
                if child.tag == f"{{{W}}}t" and (child.text or "").strip():
                    run_has_content = True
                    break
                # Drawing/image in run
                if child.tag == f"{{{W}}}drawing":
                    run_has_content = True
                    break
        if run_has_content:
            continue
        # Empty paragraph — remove if it has no meaningful properties
        ppr = p.find(f"{{{W}}}pPr")
        if ppr is None or len(ppr) == 0:
            body.remove(p)
            stats["emptyPara"] += 1

    # --- 7. Cap excessive spacing ---
    MAX_SPACING = 400  # twips (~7mm); normal paragraph spacing is 60-240
    for spacing in body.findall(f".//{{{W}}}spacing"):
        for attr in ["before", "after"]:
            val = spacing.get(f"{{{W}}}{attr}")
            if val and val.isdigit() and int(val) > MAX_SPACING:
                spacing.set(f"{{{W}}}{attr}", str(MAX_SPACING))
                stats["spacingCapped"] += 1

    # --- Write back to DOCX zip ---
    new_xml = etree.tostring(root, xml_declaration=True,
                              encoding="UTF-8", standalone=True)

    tmp_path = str(docx_path) + ".tmp"
    with zipfile.ZipFile(str(docx_path), "r") as zin:
        with zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_path, str(docx_path))
    logger.debug(f"pdf2docx cleanup: {stats}")


# -- Public API ---------------------------------------------------------------

def compare_pdfs(
    old_path: Path,
    new_path: Path,
    output_pdf: Path,
    original_name: str | None = None,
    modified_name: str | None = None,
):
    """Compare two PDFs via DOCX conversion + OOXML track-changes engine.

    Pipeline:
      1. Convert both PDFs to DOCX (pdf2docx)
      2. Clean up per-page section breaks
      3. Run OOXML comparison (same engine as .docx files)
      4. Export to PDF via Word headless
      5. Append summary/legend page

    Returns summary dict with word counts.
    """
    from doccompare.comparison.ooxml_engine import compare as ooxml_compare
    from doccompare.rendering.pdf_pipeline import produce_pdf

    old_path = Path(old_path)
    new_path = Path(new_path)
    output_pdf = Path(output_pdf)

    original_name = original_name or old_path.name
    modified_name = modified_name or new_path.name

    # Step 1-2: Convert PDFs to DOCX
    old_docx = _pdf_to_docx(old_path)
    new_docx = _pdf_to_docx(new_path)

    try:
        # Step 3: OOXML comparison (same engine as .docx files)
        import tempfile
        result_tmp = tempfile.NamedTemporaryFile(
            suffix=".docx", delete=False, prefix="doccompare_result_"
        )
        result_tmp.close()
        result_docx = Path(result_tmp.name)

        logger.info("Running OOXML comparison on converted documents")
        doc_tree, summary = ooxml_compare(
            str(old_docx), str(new_docx), str(result_docx),
        )

        # Step 4-5: Export to PDF via Word + summary page
        logger.info("Producing PDF via Word headless export")
        produce_pdf(
            doc_tree, output_pdf, summary,
            original_name, modified_name,
            docx_path=str(result_docx),
        )

        logger.info(f"PDF comparison report saved: {output_pdf}")
        return summary

    finally:
        # Clean up temp files
        old_docx.unlink(missing_ok=True)
        new_docx.unlink(missing_ok=True)
        if "result_docx" in locals():
            result_docx.unlink(missing_ok=True)
