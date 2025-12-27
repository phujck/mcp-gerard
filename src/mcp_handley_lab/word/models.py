"""Pydantic models for Word document MCP tool."""

from pydantic import BaseModel, Field


class Block(BaseModel):
    """A block-level element in a Word document (paragraph, heading, or table)."""

    id: str
    type: str  # "paragraph", "heading1".."heading9", or "table"
    text: str  # Content (for tables: markdown preview)
    style: str  # Word style name
    level: int = 0  # Heading level (1-9) or 0 for non-headings
    rows: int = 0  # Row count (tables only)
    cols: int = 0  # Column count (tables only)


class DocumentMeta(BaseModel):
    """Document metadata from core properties."""

    title: str = ""
    author: str = ""
    created: str = ""
    modified: str = ""
    revision: int = 0
    sections: int = 0


class CellInfo(BaseModel):
    """A table cell with row/column coordinates."""

    row: int  # 1-based row number
    col: int  # 1-based column number
    text: str  # Cell text content


class RunInfo(BaseModel):
    """A run (text segment with uniform formatting) within a paragraph."""

    index: int  # 0-based position in paragraph
    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font_name: str | None = None
    font_size: float | None = None  # in points
    color: str | None = None  # hex color (e.g., "FF0000")
    has_complex_content: bool = False  # True if run contains drawings/page breaks


class CommentInfo(BaseModel):
    """A comment in the document."""

    id: int  # python-docx comment_id
    author: str
    initials: str | None = None
    timestamp: str | None = None  # ISO format
    text: str


class HeaderFooterInfo(BaseModel):
    """Header and footer info for a document section."""

    section_index: int  # 0-based section index
    header_text: str | None = None  # None if linked to previous
    footer_text: str | None = None  # None if linked to previous
    header_is_linked: bool = True  # True = inherits from previous section
    footer_is_linked: bool = True
    first_page_header_text: str | None = None  # Only if different_first_page
    first_page_footer_text: str | None = None
    has_different_first_page: bool = False


class PageSetupInfo(BaseModel):
    """Page setup info for a document section."""

    section_index: int  # 0-based section index
    orientation: str  # "portrait" or "landscape"
    page_width: float  # inches
    page_height: float  # inches
    top_margin: float  # inches
    bottom_margin: float  # inches
    left_margin: float  # inches
    right_margin: float  # inches


class DocumentReadResult(BaseModel):
    """Result from read() tool."""

    block_count: int  # Count of matched blocks
    blocks: list[Block] = Field(default_factory=list)
    meta: DocumentMeta | None = None  # Only populated for scope='meta'
    cells: list[CellInfo] = Field(default_factory=list)  # For scope='table_cells'
    table_rows: int = 0  # For scope='table_cells'
    table_cols: int = 0  # For scope='table_cells'
    runs: list[RunInfo] = Field(default_factory=list)  # For scope='runs'
    comments: list[CommentInfo] = Field(default_factory=list)  # For scope='comments'
    headers_footers: list[HeaderFooterInfo] = Field(
        default_factory=list
    )  # For scope='headers_footers'
    page_setup: list[PageSetupInfo] = Field(
        default_factory=list
    )  # For scope='page_setup'
    warnings: list[str] = Field(default_factory=list)


class EditResult(BaseModel):
    """Result from edit() tool."""

    success: bool
    element_id: str = ""  # ID of created/modified block
    comment_id: int | None = None  # ID of created comment (for add_comment)
    message: str  # Human-readable status
    warnings: list[str] = Field(default_factory=list)
