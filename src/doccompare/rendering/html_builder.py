import html
from datetime import datetime
from pathlib import Path
from doccompare.models import ComparisonResult, DiffElement, DiffSegment, DiffType, ElementType

_CSS_PATH = Path(__file__).parent / "styles.css"


CSS_CLASSES = {
    DiffType.ADDED: "added",
    DiffType.DELETED: "deleted",
    DiffType.MOVED_FROM: "moved-from",
    DiffType.MOVED_TO: "moved-to",
    DiffType.UNCHANGED: "",
    DiffType.MODIFIED: "",
}


class HtmlBuilder:
    def build(
        self,
        result: ComparisonResult,
        original_path: Path,
        modified_path: Path,
    ) -> str:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        summary = result.summary

        header = f"""
        <div class="doc-header">
            <h1>Dokumentjämförelse</h1>
            <div class="meta">
                <strong>Original:</strong> {html.escape(original_path.name)} &nbsp;|&nbsp;
                <strong>Modifierat:</strong> {html.escape(modified_path.name)} &nbsp;|&nbsp;
                <strong>Datum:</strong> {now}
            </div>
        </div>
        """

        summary_box = f"""
        <div class="summary-box">
            <span class="stat stat-added">&#43; {summary.get('added_words', 0)} ord tillagda</span>
            <span class="stat stat-deleted">&#8722; {summary.get('deleted_words', 0)} ord borttagna</span>
            <span class="stat stat-moved">&#8644; {summary.get('moved_words', 0)} ord flyttade</span>
            <span class="stat">{summary.get('unchanged_words', 0)} ord oförändrade</span>
        </div>
        """

        legend = """
        <div class="summary-box">
            <strong>Legend:</strong> &nbsp;
            <span class="added">Tillagd text</span> &nbsp;&nbsp;
            <span class="deleted">Borttagen text</span> &nbsp;&nbsp;
            <span class="moved-to">Flyttad text (ny plats)</span> &nbsp;&nbsp;
            <span class="moved-from">Flyttad text (ursprunglig plats)</span>
        </div>
        """

        body_parts = [header]

        for elem in result.diff_elements:
            body_parts.append(self._render_element(elem))

        body_parts.append(summary_box)
        body_parts.append(legend)

        body_html = "\n".join(body_parts)

        css = _CSS_PATH.read_text(encoding="utf-8")
        return f"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="UTF-8">
<title>DocCompare — {html.escape(original_path.name)} vs {html.escape(modified_path.name)}</title>
<style>{css}</style>
</head>
<body>
{body_html}
</body>
</html>"""

    def _render_element(self, elem: DiffElement) -> str:
        elem_class = ""
        if elem.diff_type == DiffType.ADDED:
            elem_class = "element-added"
        elif elem.diff_type == DiffType.DELETED:
            elem_class = "element-deleted"

        inner = self._render_segments(elem.segments)

        if elem.element_type == ElementType.HEADING:
            level = max(1, min(6, elem.level or 1))
            return f'<h{level} class="{elem_class}">{inner}</h{level}>'
        elif elem.element_type == ElementType.LIST_ITEM:
            return f'<p class="list-item {elem_class}">&#8226; {inner}</p>'
        elif elem.element_type == ElementType.TABLE_ROW:
            cells_html = "".join(
                f"<td>{html.escape(c.plain_text)}</td>"
                for c in elem.segments  # reusing segments field for table rows is unusual; handle gracefully
            )
            return f'<tr class="{elem_class}">{cells_html}</tr>'
        elif elem.element_type == ElementType.PAGE_BREAK:
            return '<hr class="page-break">'
        else:
            if not inner.strip():
                return ""
            return f'<p class="{elem_class}">{inner}</p>'

    def _render_segments(self, segments: list) -> str:
        from doccompare.models import TextFormatting
        parts = []
        for seg in segments:
            escaped = html.escape(seg.text)
            # Apply formatting from the modified document
            fmt = seg.original_formatting
            if TextFormatting.BOLD in fmt:
                escaped = f"<strong>{escaped}</strong>"
            if TextFormatting.ITALIC in fmt:
                escaped = f"<em>{escaped}</em>"
            if TextFormatting.UNDERLINE in fmt:
                escaped = f"<u>{escaped}</u>"
            css = CSS_CLASSES.get(seg.diff_type, "")
            if css:
                parts.append(f'<span class="{css}">{escaped}</span>')
            else:
                parts.append(escaped)
        return "".join(parts)
