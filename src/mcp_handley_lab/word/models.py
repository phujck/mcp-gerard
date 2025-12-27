"""Pydantic models for Word document MCP tool."""

from pydantic import BaseModel, Field


class Block(BaseModel):
    """A block-level element in a Word document (paragraph, heading, or table)."""

    id: str
    type: str  # "paragraph", "heading", "table"
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


class DocumentReadResult(BaseModel):
    """Result from read() tool."""

    version: str  # SHA256 hash of document.xml for concurrency check
    block_count: int  # Total blocks in document
    blocks: list[Block] = Field(
        default_factory=list
    )  # Ordered list of paragraphs/tables
    meta: DocumentMeta | None = None  # Only populated for scope='meta'
    warnings: list[str] = Field(default_factory=list)  # TOC fields, protection, etc.


class EditResult(BaseModel):
    """Result from edit() tool."""

    success: bool
    new_version: str  # Updated doc version after edit
    element_id: str = ""  # ID of created/modified block
    message: str  # Human-readable status
    warnings: list[str] = Field(default_factory=list)
