"""Page listing operations for Visio."""

from __future__ import annotations

from mcp_handley_lab.microsoft.visio.constants import find_v, findall_v
from mcp_handley_lab.microsoft.visio.models import PageInfo
from mcp_handley_lab.microsoft.visio.ops.core import get_cell_float
from mcp_handley_lab.microsoft.visio.package import VisioPackage


def list_pages(pkg: VisioPackage) -> list[PageInfo]:
    """List all pages in the Visio document.

    Returns PageInfo for each page with name, dimensions, shape count,
    and background flag.
    """
    pages_xml = pkg.get_pages_xml()
    if pages_xml is None:
        return []

    page_els = findall_v(pages_xml, "Page")
    page_paths = pkg.get_page_paths()
    path_by_num = {num: partname for num, _rid, partname in page_paths}

    results = []
    for idx, page_el in enumerate(page_els, start=1):
        name = page_el.get("Name") or page_el.get("NameU")

        # Background flag
        background = page_el.get("Background") == "1"

        # Get page dimensions from PageSheet if available
        width = None
        height = None
        page_sheet = find_v(page_el, "PageSheet")
        if page_sheet is not None:
            width = get_cell_float(page_sheet, "PageWidth")
            height = get_cell_float(page_sheet, "PageHeight")

        # Count shapes on this page
        shape_count = 0
        if idx in path_by_num:
            try:
                page_xml = pkg.get_xml(path_by_num[idx])
                shapes = findall_v(page_xml, "Shapes")
                if shapes:
                    shape_count = len(findall_v(shapes[0], "Shape"))
            except (KeyError, IndexError):
                pass

        results.append(
            PageInfo(
                number=idx,
                name=name,
                width_inches=width,
                height_inches=height,
                shape_count=shape_count,
                is_background=background,
            )
        )

    return results


def get_page_dimensions(pkg: VisioPackage, page_num: int) -> tuple[float, float]:
    """Get page dimensions (width, height) in inches.

    Returns (width, height) or (8.5, 11.0) as default if not found.
    """
    pages_xml = pkg.get_pages_xml()
    if pages_xml is None:
        return (8.5, 11.0)

    page_els = findall_v(pages_xml, "Page")
    if page_num < 1 or page_num > len(page_els):
        return (8.5, 11.0)

    page_el = page_els[page_num - 1]
    page_sheet = find_v(page_el, "PageSheet")
    if page_sheet is None:
        return (8.5, 11.0)

    width = get_cell_float(page_sheet, "PageWidth") or 8.5
    height = get_cell_float(page_sheet, "PageHeight") or 11.0
    return (width, height)
