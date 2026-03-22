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


# -- PDF-native diff rendering ------------------------------------------------
#
# Strategy: clone the new PDF, mark additions with blue underline,
# insert deleted text with red strikethrough.
# Original formatting is 100% preserved.

_BLUE = (0.18, 0.59, 0.83)    # #2e97d3 -- additions (matches Word)
_RED = (0.71, 0.03, 0.18)     # #b5082e -- deletions (matches Word)


def _extract_fonts(doc) -> dict:
    """Extract embedded fonts from the PDF for text rewriting."""
    import fitz

    font_map = {}
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        for f in page.get_fonts(full=True):
            xref = f[0]
            name = f[3]
            base = name.split("+")[-1] if "+" in name else name
            if base in font_map:
                continue
            try:
                buf = doc.extract_font(xref)[-1]
                if buf:
                    font_map[base] = fitz.Font(fontbuffer=buf)
            except Exception:
                pass
    return font_map


def _get_font(font_map: dict, font_name: str):
    """Look up a font, falling back to base-name matching then Helvetica."""
    import fitz

    base = font_name.split("+")[-1] if "+" in font_name else font_name
    if base in font_map:
        return font_map[base]

    for name, font in font_map.items():
        if name.lower() == base.lower():
            return font

    try:
        return fitz.Font("helv")
    except Exception:
        return None


def _mark_text(doc, text: str, color: tuple, underline=False,
               strikethrough=False, start_page=0):
    """Search for text in the PDF and add colored annotation."""
    import fitz

    if not text or len(text.strip()) < 2:
        return

    search = text.strip()
    # For very long text, split into chunks
    if len(search) > 120:
        _mark_text(doc, search[:80], color, underline, strikethrough, start_page)
        _mark_text(doc, search[80:], color, underline, strikethrough, start_page)
        return

    # Try nearby pages first, then all pages
    page_ranges = [
        range(start_page, min(start_page + 3, len(doc))),
        range(len(doc)),
    ]

    for page_range in page_ranges:
        for pi in page_range:
            page = doc[pi]
            quads = page.search_for(search, quads=True)
            if quads:
                if underline:
                    annot = page.add_underline_annot(quads)
                    annot.set_colors(stroke=list(color))
                    annot.update()
                if strikethrough:
                    annot = page.add_strikeout_annot(quads)
                    annot.set_colors(stroke=list(color))
                    annot.update()
                return


def _insert_deleted_text(doc, deleted_text: str, page_num: int,
                         segments: list, font_map: dict):
    """Insert deleted text as a FreeText annotation near the paragraph."""
    import fitz

    if not deleted_text.strip() or len(deleted_text.strip()) < 2:
        return

    page = doc[min(page_num, len(doc) - 1)]

    # Find context: nearest equal/added text to anchor the position
    anchor_rect = None
    for seg_type, seg_text in segments:
        if seg_type in ("equal", "added") and len(seg_text.strip()) >= 4:
            search = seg_text.strip()[:60]
            rects = page.search_for(search)
            if rects:
                anchor_rect = rects[0]
                break

    if not anchor_rect:
        return

    # Get font info from anchor area
    fontsize = 10.1
    font = None
    td = page.get_text("dict", clip=anchor_rect)
    for block in td.get("blocks", []):
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                fontsize = span["size"]
                font = _get_font(font_map, span["font"])
                break
            if font:
                break
        if font:
            break

    if not font:
        font = _get_font(font_map, "ArialMT")
    if not font:
        return

    text_to_insert = deleted_text.strip()
    if len(text_to_insert) > 150:
        text_to_insert = text_to_insert[:150] + "..."

    # Calculate width needed
    text_length = font.text_length(text_to_insert, fontsize=fontsize)

    # Position above the anchor line
    insert_rect = fitz.Rect(
        anchor_rect.x0,
        anchor_rect.y0 - fontsize - 2,
        min(anchor_rect.x0 + text_length + 4, page.rect.width - 20),
        anchor_rect.y0 - 1,
    )

    # Ensure rect is valid and on page
    if insert_rect.y0 < 10:
        insert_rect.y0 = anchor_rect.y1 + 1
        insert_rect.y1 = anchor_rect.y1 + fontsize + 3

    # Add as FreeText annotation with red text
    try:
        annot = page.add_freetext_annot(
            insert_rect,
            text_to_insert,
            fontsize=fontsize * 0.85,
            fontname="helv",
            text_color=_RED,
            border_color=None,
            fill_color=(1, 1, 1),
        )
        annot.set_opacity(0.9)
        annot.update()
    except Exception as e:
        logger.debug(f"Could not insert deleted text annotation: {e}")


def _is_numbering_only(text: str) -> bool:
    """Check if text is only a section number like '1.12 ' or '(a) '."""
    return bool(_RE_LEADING_NUM.fullmatch(text)) or bool(
        re.fullmatch(r"\d+(?:\.\d+)*\.?\s*", text)
    )


def _only_numbering_changed(segments: list[tuple[str, str]]) -> bool:
    """Check if the only changes in a paragraph are numbering changes."""
    for seg_type, seg_text in segments:
        if seg_type in ("added", "deleted"):
            if not _is_numbering_only(seg_text):
                return False
    return True


def _apply_native_diff(doc, old_paras, new_paras, font_map):
    """Apply diff colors to the new PDF document."""
    old_texts = [p["text"] for p in old_paras]
    new_texts = [p["text"] for p in new_paras]

    matches = _match_paragraphs(old_texts, new_texts)
    matched_old = {i for i, _ in matches}
    matched_new = {j for _, j in matches}

    # Process each matched pair
    for oi, nj in matches:
        segments = _diff_paragraph(old_paras[oi]["text"], new_paras[nj]["text"])

        has_changes = any(t != "equal" for t, _ in segments)
        if not has_changes:
            continue

        # Skip paragraphs where only the numbering changed (renumbering noise)
        if _only_numbering_changed(segments):
            continue

        page_hint = new_paras[nj].get("page_num", 0)

        for seg_type, seg_text in segments:
            # Skip numbering-only segments (e.g. "1.11 " -> "1.12 ")
            if _is_numbering_only(seg_text):
                continue

            if seg_type == "added" and seg_text.strip():
                _mark_text(doc, seg_text, _BLUE, underline=True,
                           start_page=page_hint)
            elif seg_type == "deleted" and seg_text.strip():
                _insert_deleted_text(doc, seg_text, page_hint,
                                     segments, font_map)

    # Entirely new paragraphs: mark all text as blue underline
    for j in range(len(new_paras)):
        if j not in matched_new:
            page_hint = new_paras[j].get("page_num", 0)
            _mark_text(doc, new_paras[j]["text"], _BLUE, underline=True,
                       start_page=page_hint)


# -- Public API ---------------------------------------------------------------

def compare_pdfs(
    old_path: Path,
    new_path: Path,
    output_pdf: Path,
    original_name: str | None = None,
    modified_name: str | None = None,
):
    """Compare two PDFs preserving original layout and formatting.

    Uses the new PDF as the visual base:
    - Added text marked with blue underline (in-place)
    - Deleted text inserted as red annotations
    - 100% of original formatting, fonts, and layout preserved
    - Summary/legend page appended at the end.

    Returns summary dict with word counts.
    """
    import fitz
    import tempfile

    old_path = Path(old_path)
    new_path = Path(new_path)
    output_pdf = Path(output_pdf)

    original_name = original_name or old_path.name
    modified_name = modified_name or new_path.name

    logger.info(f"Extracting text from {old_path.name}")
    old_paras = _extract_paragraphs(old_path)
    logger.info(f"Extracted {len(old_paras)} paragraphs from original")

    logger.info(f"Extracting text from {new_path.name}")
    new_paras = _extract_paragraphs(new_path)
    logger.info(f"Extracted {len(new_paras)} paragraphs from modified")

    logger.info("Computing diff")
    summary = _compute_summary(old_paras, new_paras)

    logger.info("Applying diff annotations to PDF")
    doc = fitz.open(str(new_path))
    font_map = _extract_fonts(doc)

    _apply_native_diff(doc, old_paras, new_paras, font_map)

    annotated_bytes = doc.tobytes()
    doc.close()

    # Collect fully deleted paragraphs for summary page
    old_texts = [p["text"] for p in old_paras]
    new_texts = [p["text"] for p in new_paras]
    matches = _match_paragraphs(old_texts, new_texts)
    matched_old = {i for i, _ in matches}
    deletions = [old_paras[i]["text"] for i in range(len(old_paras))
                 if i not in matched_old]

    # Generate summary page and merge
    from doccompare.rendering.pdf_pipeline import _render_summary_pdf, _merge_pdfs

    summary_bytes = _render_summary_pdf(
        summary, original_name, modified_name, deletions=deletions,
    )

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(annotated_bytes)
        tmp_path = Path(tmp.name)

    try:
        _merge_pdfs(tmp_path, summary_bytes, output_pdf)
    finally:
        tmp_path.unlink(missing_ok=True)

    logger.info(f"PDF comparison report saved: {output_pdf}")
    return summary
