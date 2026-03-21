from pathlib import Path
from lxml import etree
import docx
import docx.text.paragraph
import docx.table
from docx.oxml.ns import qn
from doccompare.models import (
    ParsedDocument, DocumentElement, TextRun, ElementType, TextFormatting
)
from .base import DocumentParser


HEADING_STYLES = {
    "Heading 1": 1, "Heading 2": 2, "Heading 3": 3,
    "Heading 4": 4, "Heading 5": 5, "Heading 6": 6,
    "heading 1": 1, "heading 2": 2, "heading 3": 3,
}

LIST_STYLES = {"List Bullet", "List Number", "List Paragraph", "List Continue"}


class DocxParser(DocumentParser):
    def supports(self, file_path: Path) -> bool:
        return file_path.suffix.lower() == ".docx"

    def parse(self, file_path: Path) -> ParsedDocument:
        doc = docx.Document(str(file_path))
        elements = []
        para_idx = 0
        tbl_idx = 0

        body = doc.element.body
        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                para = docx.text.paragraph.Paragraph(child, doc)
                elem = self._parse_paragraph(para, f"p_{para_idx}", doc)
                if elem is not None:
                    elements.append(elem)
                para_idx += 1
            elif tag == "tbl":
                tbl = docx.table.Table(child, doc)
                rows = self._parse_table(tbl, f"t_{tbl_idx}")
                elements.extend(rows)
                tbl_idx += 1

        metadata = {}
        if doc.core_properties.title:
            metadata["title"] = doc.core_properties.title
        if doc.core_properties.author:
            metadata["author"] = doc.core_properties.author

        return ParsedDocument(elements=elements, metadata=metadata)

    def _get_list_info(self, doc, num_id: int, ilvl: int) -> dict:
        """Return num_fmt, lvl_text, and start_val for a numbering level."""
        result = {'num_fmt': 'bullet', 'lvl_text': '', 'start_val': 1}
        try:
            np = doc.part.numbering_part
            if np is None:
                return result
            nxml = np._element
            for num in nxml.findall(qn('w:num')):
                if num.get(qn('w:numId')) == str(num_id):
                    abid_el = num.find(qn('w:abstractNumId'))
                    if abid_el is None:
                        continue
                    abid = abid_el.get(qn('w:val'))
                    for ab in nxml.findall(qn('w:abstractNum')):
                        if ab.get(qn('w:abstractNumId')) == abid:
                            for lvl in ab.findall(qn('w:lvl')):
                                if lvl.get(qn('w:ilvl')) == str(ilvl):
                                    nf = lvl.find(qn('w:numFmt'))
                                    lt = lvl.find(qn('w:lvlText'))
                                    sv = lvl.find(qn('w:start'))
                                    if nf is not None:
                                        result['num_fmt'] = nf.get(qn('w:val'), 'bullet')
                                    if lt is not None:
                                        result['lvl_text'] = lt.get(qn('w:val'), '')
                                    if sv is not None:
                                        result['start_val'] = int(sv.get(qn('w:val'), '1'))
                                    return result
        except Exception:
            pass
        return result

    def _parse_paragraph(self, para, element_id: str, doc=None) -> "DocumentElement | None":
        style_name = para.style.name if para.style else "Normal"

        # Determine element type
        list_style = ""
        list_numid = 0
        list_lvl_text = ""
        if style_name in HEADING_STYLES:
            elem_type = ElementType.HEADING
            level = HEADING_STYLES[style_name]
        else:
            # Check for DOCX numbering via w:numPr
            pPr = para._p.find(qn('w:pPr'))
            numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
            if numPr is not None:
                ilvl_el = numPr.find(qn('w:ilvl'))
                numId_el = numPr.find(qn('w:numId'))
                if ilvl_el is not None and numId_el is not None:
                    ilvl = int(ilvl_el.get(qn('w:val'), 0))
                    numid = int(numId_el.get(qn('w:val'), 0))
                    if numid != 0:
                        elem_type = ElementType.LIST_ITEM
                        level = ilvl
                        list_numid = numid
                        info = self._get_list_info(doc, numid, ilvl) if doc is not None else {}
                        list_style = info.get('num_fmt', 'bullet')
                        list_lvl_text = info.get('lvl_text', '')
                    else:
                        elem_type = ElementType.PARAGRAPH
                        level = 0
                else:
                    elem_type = ElementType.PARAGRAPH
                    level = 0
            elif any(s in style_name for s in LIST_STYLES) or "List" in style_name:
                elem_type = ElementType.LIST_ITEM
                level = 0
            else:
                elem_type = ElementType.PARAGRAPH
                level = 0

        runs = []
        for run in para.runs:
            if not run.text:
                continue
            fmt = set()
            if run.bold:
                fmt.add(TextFormatting.BOLD)
            if run.italic:
                fmt.add(TextFormatting.ITALIC)
            if run.underline:
                fmt.add(TextFormatting.UNDERLINE)
            if run.font.strike:
                fmt.add(TextFormatting.STRIKETHROUGH)
            runs.append(TextRun(
                text=run.text,
                formatting=fmt,
                font_name=run.font.name,
                font_size=float(run.font.size.pt) if run.font.size else None,
            ))

        # Skip completely empty paragraphs with no style significance
        if not runs and elem_type == ElementType.PARAGRAPH:
            return None

        # Extract paragraph-level formatting
        pf = para.paragraph_format
        alignment = None
        if pf.alignment is not None:
            _align_map = {0: "left", 1: "center", 2: "right", 3: "justify"}
            alignment = _align_map.get(pf.alignment, None)

        def _to_pt(length):
            """Convert a docx Length to points, or None."""
            if length is None:
                return None
            try:
                return float(length.pt)
            except Exception:
                return None

        return DocumentElement(
            element_type=elem_type,
            runs=runs,
            level=level,
            element_id=element_id,
            list_style=list_style,
            list_numid=list_numid,
            list_lvl_text=list_lvl_text,
            alignment=alignment,
            left_indent_pt=_to_pt(pf.left_indent),
            right_indent_pt=_to_pt(pf.right_indent),
            first_line_indent_pt=_to_pt(pf.first_line_indent),
            space_before_pt=_to_pt(pf.space_before),
            space_after_pt=_to_pt(pf.space_after),
            line_spacing=float(pf.line_spacing) if pf.line_spacing is not None else None,
        )

    def _parse_table(self, table, table_id: str) -> list:
        rows = []
        for r_idx, row in enumerate(table.rows):
            cells = []
            for c_idx, cell in enumerate(row.cells):
                cell_runs = []
                for para in cell.paragraphs:
                    for run in para.runs:
                        if not run.text:
                            continue
                        fmt = set()
                        if run.bold:
                            fmt.add(TextFormatting.BOLD)
                        if run.italic:
                            fmt.add(TextFormatting.ITALIC)
                        cell_runs.append(TextRun(text=run.text, formatting=fmt))
                cell_elem = DocumentElement(
                    element_type=ElementType.TABLE_CELL,
                    runs=cell_runs,
                    element_id=f"{table_id}_r{r_idx}_c{c_idx}",
                )
                cells.append(cell_elem)
            row_elem = DocumentElement(
                element_type=ElementType.TABLE_ROW,
                element_id=f"{table_id}_r{r_idx}",
                children=cells,
            )
            rows.append(row_elem)
        return rows
