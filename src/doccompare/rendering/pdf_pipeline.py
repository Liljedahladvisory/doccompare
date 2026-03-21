"""PDF pipeline: render OOXML Track-Changes tree directly to PDF.

No dependency on Microsoft Word — walks the modified XML tree and produces
HTML with inline styles, then renders via WeasyPrint.
"""

import html as html_mod
from datetime import datetime
from pathlib import Path

from lxml import etree

# ── OOXML namespaces & tags ─────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def _qn(t):
    p, l = t.split(":")
    ns = {"w": W, "xml": XML_NS}[p]
    return f"{{{ns}}}{l}"


W_P = _qn("w:p")
W_R = _qn("w:r")
W_T = _qn("w:t")
W_RPR = _qn("w:rPr")
W_PPR = _qn("w:pPr")
W_TBL = _qn("w:tbl")
W_TR = _qn("w:tr")
W_TC = _qn("w:tc")
W_INS = _qn("w:ins")
W_DEL = _qn("w:del")
W_DEL_TEXT = _qn("w:delText")
W_BODY = _qn("w:body")
W_SECT_PR = _qn("w:sectPr")
W_HYPERLINK = _qn("w:hyperlink")


# ── Run formatting extraction ───────────────────────────────────────────
def _run_style(rpr) -> str:
    """Build inline CSS from a w:rPr element."""
    if rpr is None:
        return ""
    parts = []

    # Bold
    b = rpr.find(f"{{{W}}}b")
    if b is not None:
        val = b.get(f"{{{W}}}val", "true")
        if val not in ("0", "false"):
            parts.append("font-weight:bold")

    # Italic
    i = rpr.find(f"{{{W}}}i")
    if i is not None:
        val = i.get(f"{{{W}}}val", "true")
        if val not in ("0", "false"):
            parts.append("font-style:italic")

    # Underline
    u = rpr.find(f"{{{W}}}u")
    if u is not None:
        val = u.get(f"{{{W}}}val", "single")
        if val != "none":
            parts.append("text-decoration:underline")

    # Strikethrough
    strike = rpr.find(f"{{{W}}}strike")
    if strike is not None:
        val = strike.get(f"{{{W}}}val", "true")
        if val not in ("0", "false"):
            parts.append("text-decoration:line-through")

    # Font size (w:sz is in half-points)
    sz = rpr.find(f"{{{W}}}sz")
    if sz is not None:
        try:
            pt = int(sz.get(f"{{{W}}}val", "0")) / 2
            if pt > 0:
                parts.append(f"font-size:{pt:.1f}pt")
        except (ValueError, TypeError):
            pass

    # Font name
    rfonts = rpr.find(f"{{{W}}}rFonts")
    if rfonts is not None:
        name = (rfonts.get(f"{{{W}}}ascii")
                or rfonts.get(f"{{{W}}}hAnsi")
                or rfonts.get(f"{{{W}}}cs"))
        if name:
            parts.append(f"font-family:'{name}', serif")

    # Color
    color = rpr.find(f"{{{W}}}color")
    if color is not None:
        val = color.get(f"{{{W}}}val", "")
        if val and val != "auto" and len(val) == 6:
            parts.append(f"color:#{val}")

    return ";".join(parts)


# ── Paragraph formatting extraction ─────────────────────────────────────
def _para_style(ppr) -> str:
    """Build inline CSS from a w:pPr element."""
    if ppr is None:
        return ""
    parts = []

    # Alignment (w:jc)
    jc = ppr.find(f"{{{W}}}jc")
    if jc is not None:
        align_map = {"left": "left", "center": "center", "right": "right",
                     "both": "justify", "justify": "justify"}
        val = jc.get(f"{{{W}}}val", "")
        css_align = align_map.get(val)
        if css_align:
            parts.append(f"text-align:{css_align}")

    # Indentation (w:ind) — values in twips (1/20 pt)
    ind = ppr.find(f"{{{W}}}ind")
    if ind is not None:
        left = ind.get(f"{{{W}}}left") or ind.get(f"{{{W}}}start")
        right = ind.get(f"{{{W}}}right") or ind.get(f"{{{W}}}end")
        hanging = ind.get(f"{{{W}}}hanging")
        first_line = ind.get(f"{{{W}}}firstLine")
        if left:
            try:
                parts.append(f"margin-left:{int(left)/20:.1f}pt")
            except ValueError:
                pass
        if right:
            try:
                parts.append(f"margin-right:{int(right)/20:.1f}pt")
            except ValueError:
                pass
        if hanging:
            try:
                parts.append(f"text-indent:-{int(hanging)/20:.1f}pt")
            except ValueError:
                pass
        elif first_line:
            try:
                parts.append(f"text-indent:{int(first_line)/20:.1f}pt")
            except ValueError:
                pass

    # Spacing (w:spacing)
    spacing = ppr.find(f"{{{W}}}spacing")
    if spacing is not None:
        before = spacing.get(f"{{{W}}}before")
        after = spacing.get(f"{{{W}}}after")
        line = spacing.get(f"{{{W}}}line")
        line_rule = spacing.get(f"{{{W}}}lineRule", "")
        if before:
            try:
                parts.append(f"margin-top:{int(before)/20:.1f}pt")
            except ValueError:
                pass
        if after:
            try:
                parts.append(f"margin-bottom:{int(after)/20:.1f}pt")
            except ValueError:
                pass
        if line:
            try:
                val = int(line)
                if line_rule == "exact" or line_rule == "atLeast":
                    parts.append(f"line-height:{val/20:.1f}pt")
                else:
                    # Proportional: 240 = single spacing
                    parts.append(f"line-height:{val/240:.2f}")
            except ValueError:
                pass

    return ";".join(parts)


# ── Element rendering ───────────────────────────────────────────────────
def _render_run(run, css_class=None):
    """Render a w:r to HTML."""
    rpr = run.find(W_RPR)
    style = _run_style(rpr)
    t = run.find(W_T)
    text = html_mod.escape(t.text or "") if t is not None else ""
    if not text:
        return ""

    classes = f' class="{css_class}"' if css_class else ""
    style_attr = f' style="{style}"' if style else ""
    return f"<span{classes}{style_attr}>{text}</span>"


def _render_del_run(run):
    """Render a deleted w:r (with w:delText) to HTML."""
    rpr = run.find(W_RPR)
    style = _run_style(rpr)
    dt = run.find(f"{{{W}}}delText")
    text = html_mod.escape(dt.text or "") if dt is not None else ""
    if not text:
        return ""

    style_attr = f' style="{style}"' if style else ""
    return f'<span class="deleted"{style_attr}>{text}</span>'


def _render_paragraph(para) -> str:
    """Render a w:p to HTML <p>."""
    ppr = para.find(W_PPR)
    style = _para_style(ppr)
    style_attr = f' style="{style}"' if style else ""

    # Check if entire paragraph is inserted or deleted (pPr/rPr/w:ins or w:del)
    para_class = ""
    if ppr is not None:
        rpr = ppr.find(W_RPR)
        if rpr is not None:
            if rpr.find(W_INS) is not None:
                para_class = " element-added"
            elif rpr.find(W_DEL) is not None:
                para_class = " element-deleted"

    content = []
    for child in para:
        if child.tag == W_R:
            content.append(_render_run(child))
        elif child.tag == W_INS:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_run(r, css_class="added"))
        elif child.tag == W_DEL:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_del_run(r))
        elif child.tag == W_HYPERLINK:
            for r in child:
                if r.tag == W_R:
                    content.append(_render_run(r))

    inner = "".join(content)
    if not inner.strip():
        return ""
    class_attr = f' class="{para_class.strip()}"' if para_class else ""
    return f"<p{class_attr}{style_attr}>{inner}</p>"


def _render_table(tbl) -> str:
    """Render a w:tbl to HTML <table>."""
    rows = []
    for tr in tbl:
        if tr.tag != W_TR:
            continue
        cells = []
        for tc in tr:
            if tc.tag != W_TC:
                continue
            cell_html = []
            for p in tc:
                if p.tag == W_P:
                    rendered = _render_paragraph(p)
                    if rendered:
                        cell_html.append(rendered)
            cells.append(f"<td>{''.join(cell_html)}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    return f"<table>{''.join(rows)}</table>"


def render_tracked_changes_html(
    doc_tree,
    summary: dict,
    original_name: str,
    modified_name: str,
) -> str:
    """Walk the OOXML tree and produce full HTML with tracked changes."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    body = doc_tree.find(W_BODY)

    # Render all block elements
    body_parts = []
    for child in body:
        if child.tag == W_P:
            rendered = _render_paragraph(child)
            if rendered:
                body_parts.append(rendered)
        elif child.tag == W_TBL:
            body_parts.append(_render_table(child))

    content_html = "\n".join(body_parts)

    # Summary + legend (on a new page at the end)
    added = summary.get("added_words", 0)
    deleted = summary.get("deleted_words", 0)
    unchanged = summary.get("unchanged_words", 0)

    summary_html = f"""
    <div class="summary-page">
        <h2 class="summary-title">DocCompare &mdash; Sammanfattning</h2>
        <div class="summary-meta">
            <strong>Original:</strong> {html_mod.escape(original_name)} &nbsp;|&nbsp;
            <strong>Modifierat:</strong> {html_mod.escape(modified_name)} &nbsp;|&nbsp;
            <strong>Datum:</strong> {now}
        </div>
        <div class="summary-stats">
            <div class="stat stat-added">+{added} ord tillagda</div>
            <div class="stat stat-deleted">&minus;{deleted} ord borttagna</div>
            <div class="stat">{unchanged} ord of&ouml;r&auml;ndrade</div>
        </div>
        <div class="summary-legend">
            <h3>Legend</h3>
            <p><span class="added">Tillagd text</span> &mdash;
               text som finns i det modifierade dokumentet men inte i originalet.</p>
            <p><span class="deleted">Borttagen text</span> &mdash;
               text som finns i originalet men inte i det modifierade dokumentet.</p>
            <p>Of&ouml;r&auml;ndrad text &mdash;
               text som &auml;r identisk i b&aring;da dokumenten.</p>
        </div>
        <div class="summary-footer">
            Genererad av DocCompare &mdash; Liljedahl Advisory AB
        </div>
    </div>
    """

    css_path = Path(__file__).parent / "styles.css"
    css = css_path.read_text(encoding="utf-8")

    return f"""<!DOCTYPE html>
<html lang="sv">
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
{summary_html}
</body>
</html>"""


def produce_pdf(
    doc_tree,
    output_pdf: Path,
    summary: dict,
    original_name: str,
    modified_name: str,
):
    """Render tracked-changes XML tree to PDF via WeasyPrint."""
    from doccompare.rendering.pdf_renderer import render_pdf

    html_content = render_tracked_changes_html(
        doc_tree, summary, original_name, modified_name,
    )
    css_path = Path(__file__).parent / "styles.css"
    render_pdf(html_content, css_path, output_pdf)
