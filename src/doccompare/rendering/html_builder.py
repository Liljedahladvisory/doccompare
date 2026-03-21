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

    @staticmethod
    def _para_style(elem: DiffElement) -> str:
        """Build an inline CSS style string from paragraph-level formatting."""
        parts = []
        if elem.alignment:
            parts.append(f"text-align:{elem.alignment}")
        if elem.left_indent_pt:
            parts.append(f"margin-left:{elem.left_indent_pt:.1f}pt")
        if elem.right_indent_pt:
            parts.append(f"margin-right:{elem.right_indent_pt:.1f}pt")
        if elem.first_line_indent_pt:
            parts.append(f"text-indent:{elem.first_line_indent_pt:.1f}pt")
        if elem.space_before_pt:
            parts.append(f"margin-top:{elem.space_before_pt:.1f}pt")
        if elem.space_after_pt:
            parts.append(f"margin-bottom:{elem.space_after_pt:.1f}pt")
        if elem.line_spacing and elem.line_spacing != 1.0:
            parts.append(f"line-height:{elem.line_spacing:.2f}")
        return "; ".join(parts)

    def _render_element(self, elem: DiffElement) -> str:
        elem_class = ""
        if elem.diff_type == DiffType.ADDED:
            elem_class = "element-added"
        elif elem.diff_type == DiffType.DELETED:
            elem_class = "element-deleted"

        inner = self._render_segments(elem.segments)
        pstyle = self._para_style(elem)
        style_attr = f' style="{pstyle}"' if pstyle else ""

        if elem.element_type == ElementType.HEADING:
            level = max(1, min(6, elem.level or 1))
            return f'<h{level} class="{elem_class}"{style_attr}>{inner}</h{level}>'
        elif elem.element_type == ElementType.LIST_ITEM:
            return f'<p class="list-item {elem_class}"{style_attr}>&#8226; {inner}</p>'
        elif elem.element_type == ElementType.TABLE_ROW:
            cells_html = "".join(
                f"<td>{html.escape(c.plain_text)}</td>"
                for c in elem.segments
            )
            return f'<tr class="{elem_class}">{cells_html}</tr>'
        elif elem.element_type == ElementType.PAGE_BREAK:
            return '<hr class="page-break">'
        else:
            if not inner.strip():
                return ""
            return f'<p class="{elem_class}"{style_attr}>{inner}</p>'

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
            # Build inline style from run-level properties
            style_parts = []
            if seg.font_size:
                style_parts.append(f"font-size:{seg.font_size:.1f}pt")
            if seg.font_name:
                style_parts.append(f"font-family:'{seg.font_name}', serif")
            style = f' style="{"; ".join(style_parts)}"' if style_parts else ""
            if css or style:
                parts.append(f'<span class="{css}"{style}>{escaped}</span>')
            else:
                parts.append(escaped)
        return "".join(parts)

    class _NumberingTracker:
        """Tracks list counters and formats numbering labels from lvl_text templates."""
        def __init__(self):
            self._counters: dict = {}  # (numId, ilvl) -> current value

        def next_label(self, num_id: int, ilvl: int, lvl_text: str, list_style: str) -> str:
            key = (num_id, ilvl)
            self._counters[key] = self._counters.get(key, 0) + 1
            # Reset child levels
            for k in list(self._counters):
                if k[0] == num_id and k[1] > ilvl:
                    del self._counters[k]
            count = self._counters[key]

            if list_style == 'bullet' or not lvl_text:
                return '\u2022'  # •
            if list_style == 'lowerLetter':
                letter = chr(ord('a') + (count - 1) % 26)
                label = lvl_text
                for i in range(9, 0, -1):
                    label = label.replace(f'%{i}', letter if i == ilvl + 1 else
                                         str(self._counters.get((num_id, i - 1), 1)))
            elif list_style == 'upperLetter':
                letter = chr(ord('A') + (count - 1) % 26)
                label = lvl_text
                for i in range(9, 0, -1):
                    label = label.replace(f'%{i}', letter if i == ilvl + 1 else
                                         str(self._counters.get((num_id, i - 1), 1)))
            elif list_style == 'lowerRoman':
                label = lvl_text
                for i in range(9, 0, -1):
                    v = self._counters.get((num_id, i - 1), count if i == ilvl + 1 else 1)
                    label = label.replace(f'%{i}', self._to_roman(v).lower())
            elif list_style == 'upperRoman':
                label = lvl_text
                for i in range(9, 0, -1):
                    v = self._counters.get((num_id, i - 1), count if i == ilvl + 1 else 1)
                    label = label.replace(f'%{i}', self._to_roman(v))
            else:  # decimal and others
                label = lvl_text
                for i in range(9, 0, -1):
                    v = self._counters.get((num_id, i - 1), count if i == ilvl + 1 else 1)
                    label = label.replace(f'%{i}', str(v))
            return label

        @staticmethod
        def _to_roman(n: int) -> str:
            vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
                    (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
            result = ''
            for v, s in vals:
                while n >= v:
                    result += s; n -= v
            return result or 'I'

    _LIST_TYPE_MAP = {
        'decimal': '1',
        'lowerLetter': 'a',
        'upperLetter': 'A',
        'lowerRoman': 'i',
        'upperRoman': 'I',
    }

    def _render_all_elements(self, diff_elements: list) -> str:
        parts = []
        tracker = self._NumberingTracker()
        in_list = False
        current_numid = None
        last_list_style = ''

        def close_list():
            nonlocal in_list, current_numid
            if in_list:
                parts.append("</ul>" if last_list_style in ('bullet', '') else "</ol>")
                in_list = False
                current_numid = None

        for elem in diff_elements:
            if elem.element_type == ElementType.LIST_ITEM:
                list_style = elem.list_style or ''
                numid = elem.list_numid

                if in_list and numid != current_numid:
                    close_list()

                if not in_list:
                    last_list_style = list_style
                    if list_style in ('bullet', ''):
                        parts.append('<ul style="list-style:none;padding-left:0">')
                    else:
                        parts.append('<ol style="list-style:none;padding-left:0">')
                    in_list = True
                    current_numid = numid
                else:
                    last_list_style = list_style

                label = tracker.next_label(numid, elem.list_ilvl, elem.list_lvl_text, list_style)
                elem_class = ""
                if elem.diff_type == DiffType.ADDED:
                    elem_class = "element-added"
                elif elem.diff_type == DiffType.DELETED:
                    elem_class = "element-deleted"

                inner = self._render_segments(elem.segments)
                # Use document indentation if available, else fallback to level-based
                indent = elem.left_indent_pt if elem.left_indent_pt else elem.level * 20
                pstyle = self._para_style(elem)
                # Override margin-left with list indent
                li_style = f"margin-left:{indent:.1f}pt;padding-left:4pt"
                if elem.space_before_pt:
                    li_style += f";margin-top:{elem.space_before_pt:.1f}pt"
                if elem.space_after_pt:
                    li_style += f";margin-bottom:{elem.space_after_pt:.1f}pt"
                if elem.alignment:
                    li_style += f";text-align:{elem.alignment}"
                class_attr = f' class="{elem_class}"' if elem_class else ''
                parts.append(
                    f'<li{class_attr} style="{li_style}">'
                    f'<span class="list-marker">{html.escape(label)}&nbsp;</span>{inner}</li>'
                )
            else:
                close_list()
                # Numbered headings: if a heading has numbering info, prepend the label
                if elem.element_type == ElementType.HEADING and elem.list_numid:
                    label = tracker.next_label(
                        elem.list_numid, elem.list_ilvl, elem.list_lvl_text, elem.list_style
                    )
                    elem_class = ""
                    if elem.diff_type == DiffType.ADDED:
                        elem_class = "element-added"
                    elif elem.diff_type == DiffType.DELETED:
                        elem_class = "element-deleted"
                    inner = self._render_segments(elem.segments)
                    pstyle = self._para_style(elem)
                    style_attr = f' style="{pstyle}"' if pstyle else ""
                    h_level = max(1, min(6, elem.level or 1))
                    parts.append(
                        f'<h{h_level} class="{elem_class}"{style_attr}>'
                        f'<span class="list-marker">{html.escape(label)}&nbsp;</span>{inner}</h{h_level}>'
                    )
                else:
                    parts.append(self._render_element(elem))

        close_list()
        return "\n".join(parts)
