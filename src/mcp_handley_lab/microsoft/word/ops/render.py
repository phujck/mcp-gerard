"""Render Word documents to PNG images or PDF using libreoffice.

This module provides thin wrappers around the common render utilities.
"""

from mcp_handley_lab.microsoft.common.render import (
    convert_to_pdf,
    render_pages_to_images,
)


def render_to_pdf(file_path: str) -> bytes:
    """Render a Word document to PDF.

    Args:
        file_path: Path to the Word document (.docx, .docm)

    Returns:
        PDF file contents as bytes
    """
    return convert_to_pdf(file_path)


def render_to_images(
    file_path: str,
    pages: list[int],
    dpi: int = 150,
) -> list[tuple[int, bytes]]:
    """Render Word document pages to PNG images.

    Args:
        file_path: Path to the Word document (.docx, .docm)
        pages: List of 1-based page numbers to render
        dpi: Resolution in dots per inch (72-300, default: 150)

    Returns:
        List of (page_number, png_bytes) tuples
    """
    return render_pages_to_images(file_path, pages, dpi=dpi)
