from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class ElementType(Enum):
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    LIST_ITEM = "list_item"
    TABLE_CELL = "table_cell"
    TABLE_ROW = "table_row"
    PAGE_BREAK = "page_break"


class TextFormatting(Enum):
    BOLD = "bold"
    ITALIC = "italic"
    UNDERLINE = "underline"
    STRIKETHROUGH = "strikethrough"


@dataclass
class TextRun:
    text: str
    formatting: set = field(default_factory=set)
    font_name: Optional[str] = None
    font_size: Optional[float] = None


@dataclass
class DocumentElement:
    element_type: ElementType
    runs: list = field(default_factory=list)
    level: int = 0
    element_id: str = ""
    children: list = field(default_factory=list)
    list_style: str = ""
    list_numid: int = 0
    list_lvl_text: str = ""
    # Paragraph-level formatting
    alignment: Optional[str] = None          # "left", "center", "right", "justify"
    left_indent_pt: Optional[float] = None   # points
    right_indent_pt: Optional[float] = None  # points
    first_line_indent_pt: Optional[float] = None  # points (negative = hanging)
    space_before_pt: Optional[float] = None  # points
    space_after_pt: Optional[float] = None   # points
    line_spacing: Optional[float] = None     # multiplier (1.0, 1.5, 2.0 etc)

    @property
    def plain_text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass
class ParsedDocument:
    elements: list = field(default_factory=list)
    metadata: dict = field(default_factory=dict)


class DiffType(Enum):
    UNCHANGED = "unchanged"
    ADDED = "added"
    DELETED = "deleted"
    MOVED_FROM = "moved_from"
    MOVED_TO = "moved_to"
    MODIFIED = "modified"


@dataclass
class DiffSegment:
    diff_type: DiffType
    text: str
    original_formatting: set = field(default_factory=set)
    move_id: Optional[str] = None
    font_size: Optional[float] = None
    font_name: Optional[str] = None


@dataclass
class DiffElement:
    element_type: ElementType
    level: int = 0
    segments: list = field(default_factory=list)
    diff_type: DiffType = DiffType.UNCHANGED
    list_style: str = ""
    list_numid: int = 0
    list_lvl_text: str = ""
    # Paragraph-level formatting (from modified document)
    alignment: Optional[str] = None
    left_indent_pt: Optional[float] = None
    right_indent_pt: Optional[float] = None
    first_line_indent_pt: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None


@dataclass
class ComparisonResult:
    diff_elements: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)
