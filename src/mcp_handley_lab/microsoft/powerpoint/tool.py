"""PowerPoint MCP tool for reading and editing .pptx files."""

from __future__ import annotations

from typing import Literal

from mcp.server.fastmcp import FastMCP

from mcp_handley_lab.microsoft.powerpoint.constants import EMU_PER_INCH
from mcp_handley_lab.microsoft.powerpoint.models import (
    PowerPointEditResult,
    PowerPointReadResult,
    PresentationMeta,
)
from mcp_handley_lab.microsoft.powerpoint.ops.notes import get_notes, set_notes
from mcp_handley_lab.microsoft.powerpoint.ops.placeholders import set_placeholder_text
from mcp_handley_lab.microsoft.powerpoint.ops.shapes import (
    add_shape,
    delete_shape,
    edit_shape,
    get_text_in_reading_order,
    list_shapes,
)
from mcp_handley_lab.microsoft.powerpoint.ops.slides import (
    add_slide,
    delete_slide,
    get_notes_count,
    get_slide_count,
    list_slides,
    reorder_slide,
)
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage

mcp = FastMCP("PowerPoint Tool")

ReadScope = Literal["meta", "slides", "shapes", "text", "notes"]
EditOperation = Literal[
    "create",
    "add_slide",
    "delete_slide",
    "reorder_slide",
    "set_placeholder",
    "set_notes",
    "add_shape",
    "edit_shape",
    "delete_shape",
]


@mcp.tool()
def read(
    file_path: str,
    scope: ReadScope = "slides",
    slide_num: int | None = None,
) -> dict:
    """Read from a PowerPoint presentation.

    Args:
        file_path: Path to .pptx file
        scope: What to read:
            - "meta": Presentation metadata (slide count, dimensions)
            - "slides": List all slides with summaries
            - "shapes": Shapes on a slide (spatially sorted for reading order)
            - "text": All text from a slide in reading order
            - "notes": Speaker notes for a slide
        slide_num: Required for shapes/text/notes scopes (1-based)

    Returns:
        PowerPointReadResult with scope-specific data
    """
    pkg = PowerPointPackage.open(file_path)

    if scope == "meta":
        return _read_meta(pkg).model_dump(exclude_none=True)

    elif scope == "slides":
        return _read_slides(pkg).model_dump(exclude_none=True)

    elif scope == "shapes":
        if slide_num is None:
            raise ValueError("slide_num required for shapes scope")
        return _read_shapes(pkg, slide_num).model_dump(exclude_none=True)

    elif scope == "text":
        if slide_num is None:
            raise ValueError("slide_num required for text scope")
        return _read_text(pkg, slide_num).model_dump(exclude_none=True)

    elif scope == "notes":
        if slide_num is None:
            raise ValueError("slide_num required for notes scope")
        return _read_notes(pkg, slide_num).model_dump(exclude_none=True)

    else:
        raise ValueError(f"Unknown scope: {scope}")


def _read_meta(pkg: PowerPointPackage) -> PowerPointReadResult:
    """Read presentation metadata."""
    width_emu, height_emu = pkg.get_slide_dimensions()

    return PowerPointReadResult(
        scope="meta",
        meta=PresentationMeta(
            slide_count=get_slide_count(pkg),
            slide_width_inches=width_emu / EMU_PER_INCH,
            slide_height_inches=height_emu / EMU_PER_INCH,
            notes_count=get_notes_count(pkg),
        ),
    )


def _read_slides(pkg: PowerPointPackage) -> PowerPointReadResult:
    """Read all slides summary."""
    return PowerPointReadResult(
        scope="slides",
        slides=list_slides(pkg),
    )


def _read_shapes(pkg: PowerPointPackage, slide_num: int) -> PowerPointReadResult:
    """Read shapes from a slide."""
    return PowerPointReadResult(
        scope="shapes",
        shapes=list_shapes(pkg, slide_num),
    )


def _read_text(pkg: PowerPointPackage, slide_num: int) -> PowerPointReadResult:
    """Read all text from a slide in reading order."""
    return PowerPointReadResult(
        scope="text",
        text=get_text_in_reading_order(pkg, slide_num),
    )


def _read_notes(pkg: PowerPointPackage, slide_num: int) -> PowerPointReadResult:
    """Read speaker notes for a slide."""
    notes = get_notes(pkg, slide_num)
    return PowerPointReadResult(
        scope="notes",
        notes=notes,
    )


@mcp.tool()
def edit(
    file_path: str,
    operation: EditOperation,
    slide_num: int | None = None,
    text: str | None = None,
    placeholder_type: str | None = None,
    placeholder_idx: int | None = None,
    layout_name: str | None = None,
    new_position: int | None = None,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    shape_key: str | None = None,
) -> dict:
    """Edit a PowerPoint presentation.

    Args:
        file_path: Path to .pptx file (created if operation is "create")
        operation: Edit operation:
            - "create": Create new presentation
            - "add_slide": Add slide with layout
            - "delete_slide": Delete slide
            - "reorder_slide": Move slide to new position
            - "set_placeholder": Set placeholder text
            - "set_notes": Set speaker notes
            - "add_shape": Add text box (requires x, y, width, height in inches)
            - "edit_shape": Edit shape text (requires shape_key from read shapes)
            - "delete_shape": Delete shape (requires shape_key)
        slide_num: Slide number (1-based) for slide operations
        text: Text content for set_placeholder/set_notes/add_shape/edit_shape
        placeholder_type: Type for set_placeholder ("title", "body", "subtitle")
        placeholder_idx: Index for set_placeholder (alternative to type)
        layout_name: Layout name for add_slide
        new_position: New position for reorder_slide
        x: X position in inches for add_shape
        y: Y position in inches for add_shape
        width: Width in inches for add_shape
        height: Height in inches for add_shape
        shape_key: Shape identifier (slide_num:shape_id) for edit_shape/delete_shape

    Returns:
        PowerPointEditResult with success status and message
    """
    if operation == "create":
        return _edit_create(file_path).model_dump(exclude_none=True)

    # Open existing file for other operations
    pkg = PowerPointPackage.open(file_path)

    try:
        if operation == "add_slide":
            result = _edit_add_slide(pkg, layout_name)

        elif operation == "delete_slide":
            if slide_num is None:
                raise ValueError("slide_num required for delete_slide")
            result = _edit_delete_slide(pkg, slide_num)

        elif operation == "reorder_slide":
            if slide_num is None:
                raise ValueError("slide_num required for reorder_slide")
            if new_position is None:
                raise ValueError("new_position required for reorder_slide")
            result = _edit_reorder_slide(pkg, slide_num, new_position)

        elif operation == "set_placeholder":
            if slide_num is None:
                raise ValueError("slide_num required for set_placeholder")
            if text is None:
                raise ValueError("text required for set_placeholder")
            result = _edit_set_placeholder(
                pkg, slide_num, text, placeholder_type, placeholder_idx
            )

        elif operation == "set_notes":
            if slide_num is None:
                raise ValueError("slide_num required for set_notes")
            if text is None:
                raise ValueError("text required for set_notes")
            result = _edit_set_notes(pkg, slide_num, text)

        elif operation == "add_shape":
            if slide_num is None:
                raise ValueError("slide_num required for add_shape")
            if x is None or y is None or width is None or height is None:
                raise ValueError("x, y, width, height required for add_shape")
            result = _edit_add_shape(pkg, slide_num, x, y, width, height, text or "")

        elif operation == "edit_shape":
            if shape_key is None:
                raise ValueError("shape_key required for edit_shape")
            if text is None:
                raise ValueError("text required for edit_shape")
            result = _edit_edit_shape(pkg, shape_key, text)

        elif operation == "delete_shape":
            if shape_key is None:
                raise ValueError("shape_key required for delete_shape")
            result = _edit_delete_shape(pkg, shape_key)

        else:
            raise ValueError(f"Unknown operation: {operation}")

        # Save changes
        pkg.save(file_path)
        return result.model_dump(exclude_none=True)

    except Exception as e:
        return PowerPointEditResult(
            success=False,
            message=str(e),
        ).model_dump(exclude_none=True)


def _edit_create(file_path: str) -> PowerPointEditResult:
    """Create a new presentation."""
    pkg = PowerPointPackage.new()
    pkg.save(file_path)

    return PowerPointEditResult(
        success=True,
        message=f"Created presentation: {file_path}",
    )


def _edit_add_slide(
    pkg: PowerPointPackage,
    layout_name: str | None,
) -> PowerPointEditResult:
    """Add a new slide."""
    new_num = add_slide(pkg, layout_name)

    return PowerPointEditResult(
        success=True,
        message=f"Added slide {new_num}",
        element_id=str(new_num),
    )


def _edit_delete_slide(
    pkg: PowerPointPackage,
    slide_num: int,
) -> PowerPointEditResult:
    """Delete a slide."""
    delete_slide(pkg, slide_num)

    return PowerPointEditResult(
        success=True,
        message=f"Deleted slide {slide_num}",
    )


def _edit_reorder_slide(
    pkg: PowerPointPackage,
    slide_num: int,
    new_position: int,
) -> PowerPointEditResult:
    """Reorder a slide."""
    reorder_slide(pkg, slide_num, new_position)

    return PowerPointEditResult(
        success=True,
        message=f"Moved slide {slide_num} to position {new_position}",
    )


def _edit_set_placeholder(
    pkg: PowerPointPackage,
    slide_num: int,
    text: str,
    placeholder_type: str | None,
    placeholder_idx: int | None,
) -> PowerPointEditResult:
    """Set placeholder text."""
    success = set_placeholder_text(
        pkg,
        slide_num,
        text,
        placeholder_type,
        placeholder_idx,
    )

    if success:
        return PowerPointEditResult(
            success=True,
            message=f"Set placeholder text on slide {slide_num}",
        )
    else:
        return PowerPointEditResult(
            success=False,
            message="Placeholder not found",
        )


def _edit_set_notes(
    pkg: PowerPointPackage,
    slide_num: int,
    text: str,
) -> PowerPointEditResult:
    """Set speaker notes."""
    set_notes(pkg, slide_num, text)

    return PowerPointEditResult(
        success=True,
        message=f"Set notes on slide {slide_num}",
    )


def _edit_add_shape(
    pkg: PowerPointPackage,
    slide_num: int,
    x: float,
    y: float,
    width: float,
    height: float,
    text: str,
) -> PowerPointEditResult:
    """Add a new text box shape."""
    shape_key = add_shape(pkg, slide_num, x, y, width, height, text)

    return PowerPointEditResult(
        success=True,
        message=f"Added shape on slide {slide_num}",
        element_id=shape_key,
    )


def _edit_edit_shape(
    pkg: PowerPointPackage,
    shape_key: str,
    text: str,
) -> PowerPointEditResult:
    """Edit shape text."""
    success = edit_shape(pkg, shape_key, text)

    if success:
        return PowerPointEditResult(
            success=True,
            message=f"Updated shape {shape_key}",
        )
    else:
        return PowerPointEditResult(
            success=False,
            message=f"Shape {shape_key} not found or not editable",
        )


def _edit_delete_shape(
    pkg: PowerPointPackage,
    shape_key: str,
) -> PowerPointEditResult:
    """Delete a shape."""
    success = delete_shape(pkg, shape_key)

    if success:
        return PowerPointEditResult(
            success=True,
            message=f"Deleted shape {shape_key}",
        )
    else:
        return PowerPointEditResult(
            success=False,
            message=f"Shape {shape_key} not found",
        )
