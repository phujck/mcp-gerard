"""Pydantic models for PowerPoint read/edit results."""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class PresentationMeta(BaseModel):
    """Presentation metadata."""

    model_config = ConfigDict(extra="forbid")

    slide_count: int
    slide_width_inches: float
    slide_height_inches: float
    notes_count: int


class SlideInfo(BaseModel):
    """Information about a slide."""

    model_config = ConfigDict(extra="forbid")

    number: int
    title: str | None = None
    shape_count: int = 0
    has_notes: bool = False
    layout_name: str | None = None


class ShapeInfo(BaseModel):
    """Information about a shape on a slide."""

    model_config = ConfigDict(extra="forbid")

    shape_key: str  # slide_num:shape_id for edit targeting
    shape_id: int  # cNvPr@id
    type: str  # "shape", "picture", "table", "chart", "group", "connector"
    name: str | None = None

    # Position in inches
    x_inches: float
    y_inches: float
    width_inches: float
    height_inches: float

    # True if position is inherited from layout/master (no local xfrm)
    position_inherited: bool = False

    # Reading order position
    z_order: int = 0
    reading_order: int = 0

    # Text content (if any)
    text: str | None = None

    # Placeholder info (if placeholder)
    placeholder_type: str | None = None
    placeholder_idx: int | None = None


class NotesInfo(BaseModel):
    """Speaker notes information."""

    model_config = ConfigDict(extra="forbid")

    slide_number: int
    text: str
    paragraph_count: int = 1


class LayoutInfo(BaseModel):
    """Information about a slide layout."""

    model_config = ConfigDict(extra="forbid")

    name: str
    type: str | None = None  # e.g., "title", "obj", "twoObj"
    placeholder_count: int = 0


class PowerPointReadResult(BaseModel):
    """Result from read() operation."""

    model_config = ConfigDict(extra="forbid", exclude_none=True)

    scope: str
    meta: PresentationMeta | None = None
    slides: list[SlideInfo] | None = None
    shapes: list[ShapeInfo] | None = None
    text: str | None = None
    notes: NotesInfo | None = None
    layouts: list[LayoutInfo] | None = None


class PowerPointEditResult(BaseModel):
    """Result from edit() operation."""

    model_config = ConfigDict(extra="forbid", exclude_none=True)

    success: bool
    message: str
    element_id: str | None = None
    affected_refs: list[str] | None = None
