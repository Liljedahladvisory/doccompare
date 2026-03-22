"""PDF-native document comparison engine.

Compares two PDF files by extracting text (via pdfplumber), diffing paragraphs
with the same weighted-LCS + diff_match_patch approach as the OOXML engine,
and producing an HTML report with tracked-changes styling rendered via WeasyPrint.

Architecture:
  1. Extract paragraphs from both PDFs (pdfplumber + PyMuPDF for font info)
  2. Clean up: filter headers/footers, merge cross-page paragraphs
  3. Strip leading numbering before matching (avoids renumbering noise)
  4. Match paragraphs via weighted LCS (same algorithm as ooxml_engine)
  5. Character-level diff on matched pairs (diff_match_patch)
  6. Build HTML with <span class="added"> / <span class="deleted"> markup
  7. Render to PDF via WeasyPrint
"""

import html as html_mod
import re
import tempfile
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path

import diff_match_patch as dmp_module
from lxml import etree
from loguru import logger
from rapidfuzz import fuzz

from doccompare.parsers.pdf_parser import PdfParser

# ── OOXML constants for cleanup ──────────────────────────────────────────
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_P = f"{{{_W}}}p"
_W_R = f"{{{_W}}}r"
_W_T = f"{{{_W}}}t"
_W_TAB = f"{{{_W}}}tab"
_W_PPR = f"{{{_W}}}pPr"
_W_TABS = f"{{{_W}}}tabs"
_W_IND = f"{{{_W}}}ind"
_W_BODY = f"{{{_W}}}body"
_XML_NS = "http://www.w3.org/XML/1998/namespace"

# ── Regex patterns ───────────────────────────────────────────────────────

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
    r"|^[-–—]\s*\d+\s*[-–—]"         # "- 21 -" style page numbers
    r"|MATTER\s+\d+\s*\|\s*\d+"      # "MATTER 3 | 21" anywhere
, re.IGNORECASE)


# ── Diff helpers ─────────────────────────────────────────────────────────

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
    """Weighted LCS matching of paragraph texts.

    Compares on numbering-stripped text to avoid renumbering noise,
    but uses original text for the actual diff.
    """
    n, m = len(old_paras), len(new_paras)

    # Strip numbering for matching to avoid renumbering creating mismatches
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
    as a single unit (avoids character-level noise on "1.11" → "1.12").
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


# ── Paragraph extraction & cleanup ───────────────────────────────────────

def _extract_paragraphs(pdf_path: Path) -> list[dict]:
    """Extract paragraphs from a PDF with text and basic formatting info.

    Includes post-processing:
    - Filter out headers/footers (repeated text across pages, doc IDs, page numbers)
    - Merge paragraphs broken across page boundaries

    Returns list of dicts: {text, font_size, is_bold, is_italic, page_num}.
    """
    import pdfplumber
    import fitz  # PyMuPDF

    fitz_doc = fitz.open(str(pdf_path))
    raw_paras = []  # list of (text, font_size, page_num)

    with pdfplumber.open(str(pdf_path)) as pdf:
        num_pages = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages):
            lines = page.extract_text_lines() or []
            if not lines:
                continue

            # Group lines into paragraphs (same logic as PdfParser)
            paragraphs = _group_lines(lines)

            for para_text, avg_size in paragraphs:
                para_text = para_text.strip()
                if not para_text:
                    continue
                raw_paras.append({
                    "text": para_text,
                    "font_size": avg_size,
                    "is_bold": False,
                    "is_italic": False,
                    "page_num": page_num,
                    "element_type": "paragraph",
                    "level": 0,
                })

    fitz_doc.close()

    # Step 1: Detect and filter headers/footers
    paras = _filter_headers_footers(raw_paras, num_pages)

    # Step 2: Merge paragraphs broken across page boundaries
    paras = _merge_cross_page(paras)

    logger.debug(f"After cleanup: {len(paras)} paragraphs (from {len(raw_paras)} raw)")
    return paras


def _group_lines(lines: list) -> list[tuple[str, float]]:
    """Group text lines into paragraphs based on vertical spacing.

    Returns [(text, avg_font_size), ...].
    """
    if not lines:
        return []

    paragraphs = []
    current_texts = []
    current_sizes = []
    prev_bottom = None

    for line in lines:
        top = line.get("top", 0)
        bottom = line.get("bottom", top + 12)
        height = bottom - top

        if prev_bottom is not None:
            gap = top - prev_bottom
            if gap > height * 0.8:  # Large gap = new paragraph
                if current_texts:
                    text = " ".join(current_texts)
                    avg_size = sum(current_sizes) / len(current_sizes) if current_sizes else 11.0
                    paragraphs.append((text, avg_size))
                current_texts = []
                current_sizes = []

        line_text = line.get("text", "")
        if line_text:
            current_texts.append(line_text)
            chars = line.get("chars", [])
            if chars:
                sizes = [c.get("size", 11) for c in chars if c.get("size")]
                if sizes:
                    current_sizes.extend(sizes)

        prev_bottom = bottom

    if current_texts:
        text = " ".join(current_texts)
        avg_size = sum(current_sizes) / len(current_sizes) if current_sizes else 11.0
        paragraphs.append((text, avg_size))

    return paragraphs


def _filter_headers_footers(paras: list[dict], num_pages: int) -> list[dict]:
    """Remove header/footer lines that repeat across pages.

    Strategy:
    1. Regex-match obvious patterns (doc IDs, page numbers, "MATTER X | Y")
    2. Find short texts that appear on many pages (>40% of pages) — likely headers/footers
    """
    if num_pages < 2:
        return paras

    # Count occurrences of short texts across different pages
    # Normalize: strip digits to catch "Page 1 of 42" / "Page 2 of 42" etc.
    text_page_sets: dict[str, set[int]] = {}
    for p in paras:
        t = p["text"].strip()
        if len(t) > 80:  # Headers/footers are short
            continue
        # Normalize: replace digits with # for pattern matching
        normalized = re.sub(r"\d+", "#", t)
        if normalized not in text_page_sets:
            text_page_sets[normalized] = set()
        text_page_sets[normalized].add(p["page_num"])

    # Patterns appearing on >40% of pages are headers/footers
    threshold = max(2, num_pages * 0.4)
    repeated_patterns = {
        pat for pat, pages in text_page_sets.items()
        if len(pages) >= threshold
    }

    filtered = []
    removed = 0
    for p in paras:
        t = p["text"].strip()

        # Regex filter
        if _RE_HEADER_FOOTER.search(t):
            removed += 1
            continue

        # Repeated pattern filter
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
    """Merge paragraphs that were split at page boundaries.

    Heuristic: if a paragraph at page end doesn't end with sentence-ending
    punctuation and the next paragraph on the next page starts with a
    lowercase letter, they're likely one paragraph split across pages.
    """
    if len(paras) < 2:
        return paras

    merged = [paras[0]]
    for i in range(1, len(paras)):
        prev = merged[-1]
        curr = paras[i]

        # Only merge across page boundaries
        if prev["page_num"] != curr["page_num"] and curr["page_num"] == prev["page_num"] + 1:
            prev_text = prev["text"].rstrip()
            curr_text = curr["text"].lstrip()

            # Merge if: previous doesn't end with terminal punctuation
            # AND current starts with lowercase (continuation)
            if (prev_text and curr_text
                    and prev_text[-1] not in ".!?:;\"'"
                    and curr_text[0].islower()):
                prev["text"] = prev_text + " " + curr_text
                # Keep the earlier page_num
                continue

        merged.append(curr)

    if len(merged) < len(paras):
        logger.debug(f"Merged {len(paras) - len(merged)} cross-page paragraph breaks")
    return merged


# ── HTML generation ──────────────────────────────────────────────────────

def _para_to_css(para: dict) -> str:
    """Generate inline CSS for a paragraph based on extracted formatting."""
    parts = []
    fs = para.get("font_size", 11.0)
    if fs and fs != 11.0:
        parts.append(f"font-size: {fs:.1f}pt")
    if para.get("is_bold"):
        parts.append("font-weight: bold")
    if para.get("is_italic"):
        parts.append("font-style: italic")
    return "; ".join(parts)


def _render_diff_html(
    old_paras: list[dict],
    new_paras: list[dict],
    summary: dict,
    original_name: str,
    modified_name: str,
) -> str:
    """Render the diff as an HTML document with tracked-changes styling."""
    old_texts = [p["text"] for p in old_paras]
    new_texts = [p["text"] for p in new_paras]

    matches = _match_paragraphs(old_texts, new_texts)
    matched_old = {i for i, _ in matches}
    matched_new = {j for _, j in matches}

    # Build an ordered list of output paragraphs
    output_parts = []
    new_to_old = {j: i for i, j in matches}
    old_output = set()
    prev_old_idx = -1

    for nj in range(len(new_paras)):
        if nj in new_to_old:
            oi = new_to_old[nj]
            # Insert deletions: unmatched old paragraphs between prev_old_idx and oi
            for k in range(prev_old_idx + 1, oi):
                if k not in matched_old:
                    css = _para_to_css(old_paras[k])
                    style_attr = f' style="{css}"' if css else ""
                    escaped = html_mod.escape(old_paras[k]["text"])
                    output_parts.append(
                        f'<p class="element-deleted"{style_attr}>'
                        f'<span class="deleted">{escaped}</span></p>'
                    )
                    old_output.add(k)
            prev_old_idx = oi

            # Render the matched pair with inline diff
            diff_segments = _diff_paragraph(old_paras[oi]["text"], new_paras[nj]["text"])
            css = _para_to_css(new_paras[nj])
            style_attr = f' style="{css}"' if css else ""

            spans = []
            for seg_type, seg_text in diff_segments:
                escaped = html_mod.escape(seg_text)
                if seg_type == "equal":
                    spans.append(escaped)
                elif seg_type == "added":
                    spans.append(f'<span class="added">{escaped}</span>')
                elif seg_type == "deleted":
                    spans.append(f'<span class="deleted">{escaped}</span>')

            output_parts.append(f'<p{style_attr}>{"".join(spans)}</p>')
        else:
            # Unmatched new paragraph = insertion
            css = _para_to_css(new_paras[nj])
            style_attr = f' style="{css}"' if css else ""
            escaped = html_mod.escape(new_paras[nj]["text"])
            output_parts.append(
                f'<p class="element-added"{style_attr}>'
                f'<span class="added">{escaped}</span></p>'
            )

    # Any remaining unmatched old paragraphs at the end
    for k in range(prev_old_idx + 1, len(old_paras)):
        if k not in matched_old and k not in old_output:
            css = _para_to_css(old_paras[k])
            style_attr = f' style="{css}"' if css else ""
            escaped = html_mod.escape(old_paras[k]["text"])
            output_parts.append(
                f'<p class="element-deleted"{style_attr}>'
                f'<span class="deleted">{escaped}</span></p>'
            )

    content_html = "\n".join(output_parts)

    # Summary page
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    added = summary.get("added_words", 0)
    deleted = summary.get("deleted_words", 0)
    unchanged = summary.get("unchanged_words", 0)

    css_path = Path(__file__).parent.parent / "rendering" / "styles.css"
    css = css_path.read_text(encoding="utf-8")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>DocCompare &mdash; {html_mod.escape(original_name)} vs {html_mod.escape(modified_name)}</title>
<style>
{css}

/* Summary page — forced onto new page */
.summary-page {{
    page-break-before: always;
    font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
}}
.summary-title {{
    font-size: 18pt;
    font-weight: 700;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-meta {{
    font-size: 9pt;
    color: #555;
    border-top: 1px solid #bdc3c7;
    padding-top: 8pt;
    margin-bottom: 16pt;
}}
.summary-stats {{
    margin-bottom: 20pt;
}}
.summary-stats .stat {{
    display: block;
    margin-bottom: 4pt;
    font-size: 11pt;
}}
.summary-stats .stat-added {{ color: #0047ab; }}
.summary-stats .stat-deleted {{ color: #c0392b; }}
.summary-legend h3 {{
    font-size: 13pt;
    color: #2c3e50;
    margin-bottom: 8pt;
}}
.summary-legend p {{
    font-size: 10pt;
    margin-bottom: 6pt;
}}
.summary-footer {{
    margin-top: 24pt;
    padding-top: 8pt;
    border-top: 1px solid #bdc3c7;
    font-size: 8pt;
    color: #888;
}}
</style>
</head>
<body>
    {content_html}

    <div class="summary-page">
        <h2 class="summary-title">DocCompare &mdash; Summary</h2>
        <div class="summary-meta">
            <strong>Original:</strong> {html_mod.escape(original_name)} &nbsp;|&nbsp;
            <strong>Modified:</strong> {html_mod.escape(modified_name)} &nbsp;|&nbsp;
            <strong>Date:</strong> {now}
        </div>
        <div class="summary-stats">
            <div class="stat stat-added">+{added} words added</div>
            <div class="stat stat-deleted">&minus;{deleted} words deleted</div>
            <div class="stat">{unchanged} words unchanged</div>
        </div>
        <div class="summary-legend">
            <h3>Legend</h3>
            <p><span class="added">Added text</span> &mdash;
               text present in the modified document but not in the original.</p>
            <p><span class="deleted">Deleted text</span> &mdash;
               text present in the original but not in the modified document.</p>
            <p>Unchanged text &mdash;
               text identical in both documents.</p>
        </div>
        <div class="summary-footer">
            Generated by DocCompare &mdash; a Liljedahl Legal Tech Tool from Liljedahl Advisory AB
        </div>
    </div>
</body>
</html>"""


# ── Summary computation ──────────────────────────────────────────────────

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

    # Matched pairs: character diff
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

    # Unmatched old = all deleted
    for i in range(len(old_paras)):
        if i not in matched_old:
            deleted_words += len(old_paras[i]["text"].split())

    # Unmatched new = all added
    for j in range(len(new_paras)):
        if j not in matched_new:
            added_words += len(new_paras[j]["text"].split())

    return {
        "added_words": added_words,
        "deleted_words": deleted_words,
        "unchanged_words": unchanged_words,
    }


# ── Direct PDF annotation ────────────────────────────────────────────────

def _annotate_pdf(
    new_pdf_path: Path,
    output_pdf: Path,
    old_paras: list[dict],
    new_paras: list[dict],
    summary: dict,
    original_name: str,
    modified_name: str,
):
    """Annotate the newer PDF directly with diff highlights.

    - Added text: blue underline highlight
    - Deleted text: red margin comments
    - Preserves 100% of original PDF formatting.
    """
    import fitz  # PyMuPDF

    doc = fitz.open(str(new_pdf_path))

    old_texts = [p["text"] for p in old_paras]
    new_texts = [p["text"] for p in new_paras]

    matches = _match_paragraphs(old_texts, new_texts)
    matched_old = {i for i, _ in matches}
    matched_new = {j for _, j in matches}

    # Highlight color — blue for additions
    BLUE = (0.0, 0.28, 0.67)  # #0047ab — matches the CSS .added color

    # Track annotations added per page for comment positioning
    page_comment_y: dict[int, float] = {}

    # Process matched pairs — find changed segments
    for oi, nj in matches:
        diff_segments = _diff_paragraph(old_paras[oi]["text"], new_paras[nj]["text"])

        has_changes = any(st != "equal" for st, _ in diff_segments)
        if not has_changes:
            continue

        # Collect added and deleted text for this paragraph
        added_parts = []
        deleted_parts = []
        for seg_type, seg_text in diff_segments:
            if seg_type == "added":
                added_parts.append(seg_text)
            elif seg_type == "deleted":
                deleted_parts.append(seg_text)

        # Highlight added text in the PDF (search in new PDF)
        for added_text in added_parts:
            # Search for the text across all pages
            _highlight_text(doc, added_text, BLUE)

        # Add deleted text as margin comments
        if deleted_parts:
            deleted_combined = " [...] ".join(deleted_parts)
            # Find where the new paragraph text appears to place comment nearby
            new_text_snippet = new_paras[nj]["text"][:60]
            _add_deletion_comment(doc, new_text_snippet, deleted_combined, page_comment_y)

    # Mark entirely new paragraphs (unmatched in new)
    for j in range(len(new_paras)):
        if j not in matched_new:
            text = new_paras[j]["text"]
            _highlight_text(doc, text, BLUE)

    # Mark entirely deleted paragraphs (unmatched in old)
    deleted_whole = []
    for i in range(len(old_paras)):
        if i not in matched_old:
            deleted_whole.append(old_paras[i]["text"])

    if deleted_whole:
        # Add a single comment on page 1 listing all fully deleted paragraphs
        for del_text in deleted_whole:
            _add_deletion_comment(doc, "", f"[DELETED] {del_text}", page_comment_y)

    # Save annotated PDF to temp, then merge with summary page
    annotated_bytes = doc.tobytes()
    doc.close()

    # Generate summary page and merge
    from doccompare.rendering.pdf_pipeline import _render_summary_pdf, _merge_pdfs

    summary_bytes = _render_summary_pdf(summary, original_name, modified_name)

    # Write annotated PDF to temp file, then merge
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(annotated_bytes)
        tmp_path = Path(tmp.name)

    try:
        _merge_pdfs(tmp_path, summary_bytes, output_pdf)
    finally:
        tmp_path.unlink(missing_ok=True)


def _highlight_text(doc, text: str, color: tuple):
    """Search for text in the PDF and add highlight annotations.

    Skips very short strings (<6 chars) to avoid false-positive matches
    on common substrings like "(c)", "1.1", etc.
    """
    import fitz

    # Skip very short strings — they match everywhere and create false positives
    if not text or len(text.strip()) < 6:
        return

    # Search for the text (or a reasonable chunk of it)
    # Use a middle section for better uniqueness (avoid leading numbering)
    search_text = text.strip()
    if len(search_text) > 100:
        search_text = search_text[:100]

    for page in doc:
        rects = page.search_for(search_text, quads=True)
        if rects:
            for quad in rects:
                annot = page.add_highlight_annot(quad)
                annot.set_colors(stroke=color)
                annot.set_opacity(0.35)
                annot.update()
            break  # Found on this page, no need to check others


def _add_deletion_comment(
    doc, near_text: str, deleted_text: str, page_comment_y: dict,
):
    """Add a text annotation (sticky note) for deleted text.

    Places the comment icon in the RIGHT MARGIN so it doesn't overlap content.
    """
    import fitz

    # Try to find the location of nearby text to position comment on same line
    target_page = 0
    target_y = None

    if near_text:
        search = near_text[:50] if len(near_text) > 50 else near_text
        for page_num, page in enumerate(doc):
            rects = page.search_for(search)
            if rects:
                target_page = page_num
                target_y = rects[0].y0
                break

    page = doc[target_page]

    if target_y is None:
        # Stack vertically in the right margin
        target_y = page_comment_y.get(target_page, 40)

    # Place icon in the right margin (outside main text area)
    margin_x = page.rect.width - 20
    target_point = fitz.Point(margin_x, target_y)

    # Track Y position so next comment on same page doesn't overlap
    page_comment_y[target_page] = target_y + 25

    # Truncate very long deletion text
    if len(deleted_text) > 500:
        deleted_text = deleted_text[:500] + "..."

    annot = page.add_text_annot(
        target_point,
        f"Deleted: {deleted_text}",
        icon="Note",
    )
    annot.set_colors(stroke=(0.8, 0.1, 0.1))  # Dark red
    annot.set_opacity(0.8)
    annot.update()


# ── Public API ───────────────────────────────────────────────────────────

def compare_pdfs(
    old_path: Path,
    new_path: Path,
    output_pdf: Path,
    original_name: str | None = None,
    modified_name: str | None = None,
):
    """Compare two PDFs and produce an annotated diff report.

    Annotates the newer PDF directly:
    - Blue highlights for added/changed text
    - Red comment annotations for deleted text
    - Preserves 100% of original formatting.
    - Appends summary/legend page at the end.

    Returns summary dict with word counts.
    """
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

    logger.info("Annotating PDF with changes")
    _annotate_pdf(
        new_path, output_pdf,
        old_paras, new_paras, summary,
        original_name, modified_name,
    )

    logger.info(f"PDF comparison report saved: {output_pdf}")
    return summary
