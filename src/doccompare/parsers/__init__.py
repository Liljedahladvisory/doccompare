from .base import DocumentParser
from .docx_parser import DocxParser
from .pdf_parser import PdfParser
from pathlib import Path


def get_parser(file_path: Path) -> DocumentParser:
    """Return the appropriate parser for the given file."""
    suffix = file_path.suffix.lower()
    if suffix == ".docx":
        return DocxParser()
    elif suffix == ".pdf":
        return PdfParser()
    else:
        raise ValueError(f"Unsupported file format: {suffix}. Supported: .docx, .pdf")
