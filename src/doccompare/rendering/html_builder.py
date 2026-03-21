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

        body_parts.append(self._render_all_elements(result.diff_elements))

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
            fmt = seg.original_formatting
            if TextFormatting.BOLD in fmt:
                escaped = f"<strong>{escaped}</strong>"
            if TextFormatting.ITALIC in fmt:
                escaped = f"<em>{escaped}</em>"
            if TextFormatting.UNDERLINE in fmt:
                escaped = f"<u>{escaped}</u>"
            css = CSS_CLASSES.get(seg.diff_type, "")
            style = f' style="font-size:{seg.font_size:.1f}pt"' if seg.font_size else ""
            if css or style:
                parts.append(f'<span class="{css}"{style}>{escaped}</span>')
            else:
                parts.append(escaped)
        return "".join(parts)

    _LIST_TYPE_MAP = {
        'decimal': '1',
        'lowerLetter': 'a',
        'upperLetter': 'A',
        'lowerRoman': 'i',
        'upperRoman': 'I',
    }

    def _render_all_elements(self, diff_elements: list) -> str:
        parts = []
        in_list = False
        current_numid = None

        def close_list():
            nonlocal in_list, current_numid
            if in_list:
                parts.append("</ul>" if _last_list_style in ('bullet', '') else "</ol>")
                in_list = False
                current_numid = None

        _last_list_style = ''

        for elem in diff_elements:
            if elem.element_type == ElementType.LIST_ITEM:
                list_style = elem.list_style or ''
                numid = elem.list_numid

                # Close existing list if numid changed
                if in_list and numid != current_numid:
                    if list_style in ('bullet', ''):
                        parts.append("</ul>")
                    else:
                        parts.append("</ol>")
                    in_list = False
                    current_numid = None

                # Open new list if not in one
                if not in_list:
                    _last_list_style = list_style
                    if list_style in ('bullet', ''):
                        parts.append("<ul>")
                    else:
                        ol_type = self._LIST_TYPE_MAP.get(list_style, '1')
                        parts.append(f'<ol type="{ol_type}">')
                    in_list = True
                    current_numid = numid
                else:
                    _last_list_style = list_style

                elem_class = ""
                if elem.diff_type == DiffType.ADDED:
                    elem_class = "element-added"
                elif elem.diff_type == DiffType.DELETED:
                    elem_class = "element-deleted"

                inner = self._render_segments(elem.segments)
                margin = elem.level * 20
                class_attr = f' class="{elem_class}"' if elem_class else ''
                parts.append(f'<li{class_attr} style="margin-left:{margin}pt">{inner}</li>')
            else:
                if in_list:
                    if _last_list_style in ('bullet', ''):
                        parts.append("</ul>")
                    else:
                        parts.append("</ol>")
                    in_list = False
                    current_numid = None

                parts.append(self._render_element(elem))

        if in_list:
            if _last_list_style in ('bullet', ''):
                parts.append("</ul>")
            else:
                parts.append("</ol>")

        return "\n".join(parts)
