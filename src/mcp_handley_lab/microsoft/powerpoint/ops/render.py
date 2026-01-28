"""Render PowerPoint presentations to PNG images or PDF using libreoffice.

This module provides thin wrappers around the common render utilities,
with PowerPoint-specific terminology (slides vs pages).
"""

from mcp_handley_lab.microsoft.common.render import (
    convert_to_pdf,
    render_pages_to_images,
)


def render_to_pdf(file_path: str) -> bytes:
    """Render a PowerPoint presentation to PDF.

    Args:
        file_path: Path to the PowerPoint file (.pptx, .pptm, .ppsx)

    Returns:
        PDF file contents as bytes
    """
    return convert_to_pdf(file_path)


def render_to_images(
    file_path: str,
    slides: list[int],
    dpi: int = 150,
) -> list[tuple[int, bytes]]:
    """Render PowerPoint slides to PNG images.

    Args:
        file_path: Path to the PowerPoint file (.pptx, .pptm, .ppsx)
        slides: List of 1-based slide numbers to render
        dpi: Resolution in dots per inch (72-300, default: 150)

    Returns:
        List of (slide_number, png_bytes) tuples

    Raises:
        ValueError: If slides list is empty, too many slides, DPI out of range,
                   or slide out of bounds
        RuntimeError: If libreoffice/pdftoppm not found or conversion fails
    """
    # Validate with PowerPoint-specific error messages
    if not slides:
        raise ValueError("slides is required")

    unique_slides = sorted(set(slides))
    if len(unique_slides) > 5:
        raise ValueError(f"max 5 slides allowed; requested {len(unique_slides)}")

    # Delegate to common implementation
    try:
        return render_pages_to_images(file_path, slides, dpi=dpi)
    except ValueError as e:
        # Translate generic "page" errors to "slide" terminology
        msg = str(e)
        if "Page " in msg and " out of bounds" in msg:
            # Extract page number and rewrite message
            import re

            match = re.search(r"Page (\d+) out of bounds", msg)
            if match:
                raise ValueError(f"Slide {match.group(1)} out of bounds") from e
        raise
