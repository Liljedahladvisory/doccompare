"""Compatibility facade for callers of the previous adapter API."""
from abc import ABC, abstractmethod
from pathlib import Path
import sys


class ComparisonAdapter(ABC):
    @abstractmethod
    def is_available(self) -> bool:
        """Whether this platform can run the layout-preserving pipeline."""

    @abstractmethod
    def compare_and_export(self, original, modified, output_pdf,
                           original_name='', modified_name='', *, author='DocCompare') -> dict:
        """Return statistics from Word's actual revisions; raise on failure."""


class MacWordAdapter(ComparisonAdapter):
    def is_available(self):
        return sys.platform == 'darwin' and Path('/Applications/Microsoft Word.app').exists()

    def compare_and_export(self, original, modified, output_pdf,
                           original_name='', modified_name='', *, author='DocCompare'):
        from .service import compare_documents
        return compare_documents(original, modified, output_pdf, author=author,
                                 original_name=original_name, modified_name=modified_name)


def get_adapter():
    adapter = MacWordAdapter()
    return adapter if adapter.is_available() else None
