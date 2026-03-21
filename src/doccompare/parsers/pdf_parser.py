from pathlib import Path
from loguru import logger
from doccompare.models import (
    ParsedDocument, DocumentElement, TextRun, ElementType, TextFormatting
)
from .base import DocumentParser


class PdfParser(DocumentParser):
    def supports(self, file_path: Path) -> bool:
        return file_path.suffix.lower() == ".pdf"

    def parse(self, file_path: Path) -> ParsedDocument:
        import pdfplumber
        import fitz  # PyMuPDF

        elements = []
        para_idx = 0

        # Get font info from PyMuPDF
        fitz_doc = fitz.open(str(file_path))
        fitz_pages = [fitz_doc[i] for i in range(len(fitz_doc))]

        # Check if PDF has text
        total_text = ""
        with pdfplumber.open(str(file_path)) as pdf:
            for page in pdf.pages:
                total_text += page.extract_text() or ""

        if not total_text.strip():
            raise ValueError(
                "Denna PDF verkar vara skannad och saknar textlager. "
                "Prova att köra OCR först."
            )

        try:
            if fitz_doc.is_encrypted:
                raise ValueError("Denna PDF är krypterad. Dekryptera den först.")
        except Exception:
            pass

        with pdfplumber.open(str(file_path)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                fitz_page = fitz_pages[page_num]
                font_info = self._get_font_info(fitz_page)

                lines = page.extract_text_lines() or []
                paragraphs = self._group_lines_into_paragraphs(lines)

                for para_text, avg_size, is_bold, is_italic in paragraphs:
                    if not para_text.strip():
                        continue

                    elem_type, level = self._classify_element(avg_size, para_text)
                    fmt = set()
                    if is_bold:
                        fmt.add(TextFormatting.BOLD)
                    if is_italic:
                        fmt.add(TextFormatting.ITALIC)

                    run = TextRun(text=para_text, formatting=fmt, font_size=avg_size)
                    elem = DocumentElement(
                        element_type=elem_type,
                        runs=[run],
                        level=level,
                        element_id=f"p_{para_idx}",
                    )
                    elements.append(elem)
                    para_idx += 1

        fitz_doc.close()
        return ParsedDocument(elements=elements, metadata={})

    def _get_font_info(self, fitz_page) -> dict:
        """Extract font info from PyMuPDF page."""
        font_map = {}
        try:
            blocks = fitz_page.get_text("dict")["blocks"]
            for block in blocks:
                if block.get("type") == 0:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if text:
                                font_map[text[:50]] = {
                                    "size": span.get("size", 11),
                                    "flags": span.get("flags", 0),
                                }
        except Exception as e:
            logger.debug(f"Could not extract font info: {e}")
        return font_map

    def _group_lines_into_paragraphs(self, lines: list) -> list:
        """Group text lines into logical paragraphs based on vertical spacing."""
        if not lines:
            return []

        paragraphs = []
        current_lines = []
        prev_bottom = None

        for line in lines:
            top = line.get("top", 0)
            bottom = line.get("bottom", top + 12)
            height = bottom - top

            if prev_bottom is not None:
                gap = top - prev_bottom
                if gap > height * 0.8:  # Large gap = new paragraph
                    if current_lines:
                        paragraphs.append(self._merge_lines(current_lines))
                    current_lines = []

            current_lines.append(line)
            prev_bottom = bottom

        if current_lines:
            paragraphs.append(self._merge_lines(current_lines))

        return paragraphs

    def _merge_lines(self, lines: list) -> tuple:
        """Merge lines into a single paragraph text with metadata."""
        text_parts = []
        sizes = []

        for line in lines:
            words = line.get("chars", [])
            if words:
                sizes.extend([c.get("size", 11) for c in words if c.get("size")])
            line_text = line.get("text", "")
            if line_text:
                text_parts.append(line_text)

        full_text = " ".join(text_parts)
        avg_size = sum(sizes) / len(sizes) if sizes else 11.0
        return full_text, avg_size, False, False

    def _classify_element(self, font_size: float, text: str) -> tuple:
        """Classify element type based on font size heuristics."""
        if font_size >= 18:
            return ElementType.HEADING, 1
        elif font_size >= 15:
            return ElementType.HEADING, 2
        elif font_size >= 13:
            return ElementType.HEADING, 3
        elif font_size >= 12:
            return ElementType.HEADING, 4
        elif text.strip().startswith(("•", "-", "*", "–")):
            return ElementType.LIST_ITEM, 0
        elif len(text) < 5 and text.strip().endswith("."):
            return ElementType.LIST_ITEM, 0
        else:
            return ElementType.PARAGRAPH, 0
