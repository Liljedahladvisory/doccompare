from abc import ABC, abstractmethod
from pathlib import Path
from doccompare.models import ParsedDocument


class DocumentParser(ABC):
    @abstractmethod
    def parse(self, file_path: Path) -> ParsedDocument:
        ...

    @abstractmethod
    def supports(self, file_path: Path) -> bool:
        ...
