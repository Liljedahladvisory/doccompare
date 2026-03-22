"""PDF-native document comparison engine.

Compares two PDF files by extracting text (via pdfplumber), diffing paragraphs
with the same weighted-LCS + diff_match_patch approach as the OOXML engine,
and producing an HTML report with tracked-changes styling rendered via WeasyPrint.

Architecture:
  1. Extract paragraphs from both PDFs (pdfplumber + PyMuPDF for font info)
  2. Match paragraphs via weighted LCS (same algorithm as ooxml_engine)
  3. Character-level diff on matched pairs (diff_match_patch)
  4. Build HTML with <span class="added"> / <span class="deleted"> markup
  5. Render to PDF via WeasyPrint, then merge summary page via pypdf
"""

import html as html_mod
import tempfile
from datetime import datetime
from pathlib import Path

import diff_match_patch as dmp_module
from loguru import logger
from rapidfuzz import fuzz

from doccompare.parsers.pdf_parser import PdfParser

# ── Diff helpers ─────────────────────────────────────────────────────────

def _similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return fuzz.ratio(a, b) / 100.0


def _match_paragraphs(old_paras: list[str], new_paras: list[str]) -> list[tuple[int, int]]:
    """Weighted LCS matching of paragraph texts."""
    n, m = len(old_paras), len(new_paras)

    sim_cache: dict = {}
    for i in range(n):
        for j in range(m):
            s = _similarity(old_paras[i], new_paras[j])
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

    type is 'equal', 'added', or 'deleted'.
    """
    dmp = dmp_module.diff_match_patch()
    diffs = dmp.diff_main(old_text, new_text)
    dmp.diff_cleanupSemantic(diffs)

    result = []
    for op, text in diffs:
        if op == 0:
            result.append(("equal", text))
        elif op == 1:
            result.append(("added", text))
        elif op == -1:
            result.append(("deleted", text))
    return result


# ── Paragraph extraction ─────────────────────────────────────────────────

def _extract_paragraphs(pdf_path: Path) -> list[dict]:
    """Extract paragraphs from a PDF with text and basic formatting info.

    Returns list of dicts: {text, font_size, is_bold, is_italic}.
    """
    parser = PdfParser()
    doc = parser.parse(pdf_path)

    paragraphs = []
    for elem in doc.elements:
        text = elem.plain_text.strip()
        if not text:
            continue
        font_size = None
        is_bold = False
        is_italic = False
        if elem.runs:
            font_size = elem.runs[0].font_size
            from doccompare.models import TextFormatting
            is_bold = TextFormatting.BOLD in elem.runs[0].formatting
            is_italic = TextFormatting.ITALIC in elem.runs[0].formatting
        paragraphs.append({
            "text": text,
            "font_size": font_size or 11.0,
            "is_bold": is_bold,
            "is_italic": is_italic,
            "element_type": elem.element_type.value,
            "level": elem.level,
        })
    return paragraphs


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
    # Walk through new paragraphs in order, inserting deleted old paragraphs
    # at the right positions
    output_parts = []

    # Create a map: for each new index, which old index is it matched to?
    new_to_old = {j: i for i, j in matches}

    # Track which old paragraphs we've output (as deletions)
    old_output = set()

    # Walk through matches to find deletion insertion points
    # Between consecutive matched pairs (oi1, nj1) and (oi2, nj2),
    # any unmatched old paragraphs in range (oi1, oi2) are deletions.
    prev_old_idx = -1

    for nj in range(len(new_paras)):
        # Before this new paragraph, insert any unmatched old paragraphs
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
            all_equal = all(seg_type == "equal" for seg_type, _ in diff_segments)
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


# ── Public API ───────────────────────────────────────────────────────────

def compare_pdfs(
    old_path: Path,
    new_path: Path,
    output_pdf: Path,
    original_name: str | None = None,
    modified_name: str | None = None,
):
    """Compare two PDFs and produce a diff report as PDF.

    Returns summary dict with word counts.
    """
    from doccompare.rendering.pdf_renderer import render_pdf

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

    logger.info("Rendering HTML")
    html_content = _render_diff_html(
        old_paras, new_paras, summary, original_name, modified_name,
    )

    logger.info("Rendering PDF via WeasyPrint")
    css_path = Path(__file__).parent.parent / "rendering" / "styles.css"
    render_pdf(html_content, css_path, output_pdf)

    logger.info(f"PDF comparison report saved: {output_pdf}")
    return summary
