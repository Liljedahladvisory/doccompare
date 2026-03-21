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
                elem = self._parse_paragraph(para, f"p_{para_idx}")
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

    def _parse_paragraph(self, para, element_id: str) -> "DocumentElement | None":
        style_name = para.style.name if para.style else "Normal"

        # Determine element type
        if style_name in HEADING_STYLES:
            elem_type = ElementType.HEADING
            level = HEADING_STYLES[style_name]
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

        return DocumentElement(
            element_type=elem_type,
            runs=runs,
            level=level,
            element_id=element_id,
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
