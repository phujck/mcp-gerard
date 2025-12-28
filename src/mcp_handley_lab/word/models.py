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
    """A table cell with coordinates and hierarchical ID."""

    row: int  # 0-based row index
    col: int  # 0-based column index
    text: str  # Cell text content
    hierarchical_id: str = ""  # e.g., "table_abc_0#r0c0"
    grid_span: int = 1  # Horizontal span (columns merged)
    row_span: int = 1  # Vertical span (rows merged)
    is_merge_origin: bool = True  # False if this is a continuation cell
    width_inches: float | None = None  # Cell width
    vertical_alignment: str | None = None  # "top", "center", "bottom"


class RowInfo(BaseModel):
    """Row properties for a table."""

    index: int  # 0-based row index
    height_inches: float | None = None
    height_rule: str | None = None  # "auto", "at_least", "exactly"


class TableLayoutInfo(BaseModel):
    """Table layout properties."""

    table_id: str
    alignment: str | None = None  # "left", "center", "right"
    autofit: bool = True
    rows: list[RowInfo] = Field(default_factory=list)


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
    highlight_color: str | None = None  # e.g., "yellow", "cyan"
    strike: bool | None = None
    double_strike: bool | None = None
    subscript: bool | None = None
    superscript: bool | None = None
    style: str | None = None  # Character style name
    is_hyperlink: bool = False  # True if inside a hyperlink
    hyperlink_url: str | None = None  # URL if inside hyperlink
    # Additional font properties
    all_caps: bool | None = None  # Text appears in capital letters
    small_caps: bool | None = None  # Lowercase as smaller capitals
    hidden: bool | None = None  # Hidden text (not displayed unless settings allow)
    emboss: bool | None = None  # Raised emboss effect
    imprint: bool | None = None  # Pressed into page effect
    outline: bool | None = None  # One-pixel border around glyphs
    shadow: bool | None = None  # Shadow effect on characters


class HyperlinkInfo(BaseModel):
    """A hyperlink within a paragraph."""

    index: int  # Position in document's hyperlink list
    text: str  # Visible link text
    url: str  # Full URL (address + fragment)
    address: str  # Base URL
    fragment: str  # Anchor/bookmark (without #)
    is_external: bool  # True if external link (has rId)


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
    even_page_header_text: str | None = None  # Only if different_odd_even
    even_page_footer_text: str | None = None
    has_different_odd_even: bool = False


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


class ImageInfo(BaseModel):
    """Information about an embedded inline image."""

    id: str  # image_{sha1[:8]}_{occurrence}
    width_inches: float
    height_inches: float
    content_type: str  # e.g., "image/png", "image/jpeg"
    block_id: (
        str  # Hierarchical ID of containing paragraph (e.g., "table_abc_0#r0c0/p0")
    )
    run_index: int  # 0-based index of run within paragraph
    image_index_in_run: int  # 0-based index among inline images in run
    filename: str  # Original filename if available


class StyleInfo(BaseModel):
    """A style definition in the document."""

    name: str  # UI name (e.g., "Heading 1")
    style_id: str  # Internal ID
    type: str  # "paragraph", "character", "table", "list"
    builtin: bool  # True if built-in style
    base_style: str | None = None  # Parent style name
    next_style: str | None = None  # Auto-applied next paragraph style
    hidden: bool = False  # Hidden from UI
    quick_style: bool = False  # In Quick Styles gallery


class StyleFormatInfo(BaseModel):
    """Detailed style formatting (returned when reading a specific style)."""

    name: str
    style_id: str
    type: str
    # Font properties (from style.font)
    font_name: str | None = None
    font_size: float | None = None  # points
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None  # hex
    # Paragraph properties (paragraph styles only)
    alignment: str | None = None
    left_indent: float | None = None  # inches
    space_before: float | None = None  # points
    space_after: float | None = None  # points
    line_spacing: float | None = None


class TabStopInfo(BaseModel):
    """A tab stop definition."""

    position_inches: float
    alignment: str  # "left", "center", "right", "decimal"
    leader: str  # "spaces", "dots", "heavy", "middle_dot"


class ParagraphFormatInfo(BaseModel):
    """Paragraph formatting properties."""

    alignment: str | None = None  # "left", "center", "right", "justify"
    left_indent: float | None = None  # inches
    right_indent: float | None = None  # inches
    first_line_indent: float | None = None  # inches
    space_before: float | None = None  # points
    space_after: float | None = None  # points
    line_spacing: float | None = None  # multiplier or points
    keep_with_next: bool | None = None
    page_break_before: bool | None = None
    tab_stops: list[TabStopInfo] = Field(default_factory=list)


class DocumentReadResult(BaseModel):
    """Result from read() tool."""

    block_count: int
    blocks: list[Block] = Field(default_factory=list)
    meta: DocumentMeta | None = None
    cells: list[CellInfo] = Field(default_factory=list)
    table_rows: int = 0
    table_cols: int = 0
    runs: list[RunInfo] = Field(default_factory=list)
    comments: list[CommentInfo] = Field(default_factory=list)
    headers_footers: list[HeaderFooterInfo] = Field(default_factory=list)
    page_setup: list[PageSetupInfo] = Field(default_factory=list)
    images: list[ImageInfo] = Field(default_factory=list)
    hyperlinks: list[HyperlinkInfo] = Field(default_factory=list)
    styles: list[StyleInfo] = Field(default_factory=list)
    paragraph_format: ParagraphFormatInfo | None = None  # For runs scope
    style_format: "StyleFormatInfo | None" = None  # For style scope
    table_layout: "TableLayoutInfo | None" = None  # For table_layout scope


class EditResult(BaseModel):
    """Result from edit() tool."""

    success: bool
    element_id: str = ""
    comment_id: int | None = None
    message: str
