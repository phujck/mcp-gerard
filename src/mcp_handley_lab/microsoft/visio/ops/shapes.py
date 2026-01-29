"""Shape extraction and text operations for Visio."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.visio.constants import findall_v
from mcp_handley_lab.microsoft.visio.models import (
    ShapeCellInfo,
    ShapeDataProperty,
    ShapeInfo,
)
from mcp_handley_lab.microsoft.visio.ops.core import (
    extract_shape_text,
    get_all_cells,
    get_cell_float,
    get_cell_value,
    get_section_rows,
    make_shape_key,
)
from mcp_handley_lab.microsoft.visio.ops.masters import resolve_master_name
from mcp_handley_lab.microsoft.visio.ops.pages import get_page_dimensions
from mcp_handley_lab.microsoft.visio.package import VisioPackage


def _get_shape_type(shape_el: etree._Element) -> str:
    """Determine shape type from Type attribute and cell presence."""
    type_attr = shape_el.get("Type")
    if type_attr == "Group":
        return "group"
    if type_attr == "Foreign":
        return "foreign"

    # Check for 1D shape (connector): has BeginX/EndX cells
    begin_x = get_cell_value(shape_el, "BeginX")
    end_x = get_cell_value(shape_el, "EndX")
    if begin_x is not None and end_x is not None:
        return "connector"

    return "shape"


def _parse_shape(
    shape_el: etree._Element,
    page_num: int,
    master_names: dict[int, str],
    parent_id: int | None = None,
) -> ShapeInfo | None:
    """Parse a single Shape element into ShapeInfo."""
    id_str = shape_el.get("ID")
    if id_str is None:
        return None
    shape_id = int(id_str)

    name = shape_el.get("Name")
    name_u = shape_el.get("NameU")
    shape_type = _get_shape_type(shape_el)

    text = extract_shape_text(shape_el)

    # Master reference
    master_id = None
    master_name = None
    master_attr = shape_el.get("Master")
    if master_attr is not None:
        master_id = int(master_attr)
        master_name = master_names.get(master_id)

    # Position: 2D shapes use PinX/PinY/Width/Height
    x = get_cell_float(shape_el, "PinX")
    y = get_cell_float(shape_el, "PinY")
    width = get_cell_float(shape_el, "Width")
    height = get_cell_float(shape_el, "Height")

    # Adjust position to top-left corner (PinX/PinY is center by default)
    loc_pin_x = get_cell_float(shape_el, "LocPinX")
    loc_pin_y = get_cell_float(shape_el, "LocPinY")
    if x is not None and loc_pin_x is not None:
        x = x - loc_pin_x
    elif x is not None and width is not None:
        x = x - width / 2
    if y is not None and loc_pin_y is not None:
        y = y - loc_pin_y
    elif y is not None and height is not None:
        y = y - height / 2

    # 1D shapes: connector endpoints
    begin_x = get_cell_float(shape_el, "BeginX")
    begin_y = get_cell_float(shape_el, "BeginY")
    end_x = get_cell_float(shape_el, "EndX")
    end_y = get_cell_float(shape_el, "EndY")

    return ShapeInfo(
        shape_id=shape_id,
        shape_key=make_shape_key(page_num, shape_id),
        name=name,
        name_u=name_u,
        type=shape_type,
        text=text,
        x_inches=x,
        y_inches=y,
        width_inches=width,
        height_inches=height,
        begin_x=begin_x,
        begin_y=begin_y,
        end_x=end_x,
        end_y=end_y,
        master_id=master_id,
        master_name=master_name,
        parent_id=parent_id,
    )


def _collect_shapes_recursive(
    parent: etree._Element,
    page_num: int,
    master_names: dict[int, str],
    parent_id: int | None = None,
) -> list[ShapeInfo]:
    """Recursively collect shapes, traversing into groups."""
    results = []

    for shape_el in findall_v(parent, "Shape"):
        info = _parse_shape(shape_el, page_num, master_names, parent_id)
        if info is None:
            continue
        results.append(info)

        # Recurse into group children
        if info.type == "group":
            shapes_container = findall_v(shape_el, "Shapes")
            for container in shapes_container:
                children = _collect_shapes_recursive(
                    container, page_num, master_names, info.shape_id
                )
                results.extend(children)

    return results


def _spatial_sort_key(shape: ShapeInfo, page_height: float) -> tuple[float, float, int]:
    """Sort key for spatial ordering: top-to-bottom, left-to-right.

    Flips Y coordinate since Visio Y is bottom-to-top.
    """
    if shape.type == "connector":
        # Use midpoint for connectors
        if shape.begin_x is not None and shape.end_x is not None:
            x = (shape.begin_x + shape.end_x) / 2
        else:
            x = 0.0
        if shape.begin_y is not None and shape.end_y is not None:
            y = page_height - (shape.begin_y + shape.end_y) / 2
        else:
            y = 0.0
    else:
        x = shape.x_inches or 0.0
        # Flip Y: page_height - y gives top-to-bottom order
        y = page_height - (shape.y_inches or 0.0)

    return (y, x, shape.shape_id)


def list_shapes(pkg: VisioPackage, page_num: int) -> list[ShapeInfo]:
    """List all shapes on a page, spatially sorted for reading order.

    Shapes are sorted top-to-bottom, left-to-right (Y-flipped for display order).
    Group children are included with parent_id set.
    """
    page_xml = pkg.get_page_xml(page_num)
    master_names = resolve_master_name(pkg)
    page_width, page_height = get_page_dimensions(pkg, page_num)

    # Find the top-level Shapes container in the page
    shapes_containers = findall_v(page_xml, "Shapes")
    if not shapes_containers:
        return []

    all_shapes = []
    for container in shapes_containers:
        all_shapes.extend(_collect_shapes_recursive(container, page_num, master_names))

    # Sort spatially
    all_shapes.sort(key=lambda s: _spatial_sort_key(s, page_height))

    # Assign reading order
    for idx, shape in enumerate(all_shapes):
        shape.reading_order = idx

    return all_shapes


def get_text_in_reading_order(pkg: VisioPackage, page_num: int) -> str:
    """Get all text from a page in spatial reading order."""
    shapes = list_shapes(pkg, page_num)
    texts = []
    for shape in shapes:
        if shape.text:
            texts.append(shape.text)
    return "\n\n".join(texts)


def find_shape_element(
    pkg: VisioPackage, page_num: int, shape_id: int
) -> etree._Element | None:
    """Find a shape element by ID on a page.

    Searches recursively into groups.
    """
    page_xml = pkg.get_page_xml(page_num)
    shapes_containers = findall_v(page_xml, "Shapes")

    for container in shapes_containers:
        found = _find_shape_recursive(container, shape_id)
        if found is not None:
            return found
    return None


def _find_shape_recursive(
    parent: etree._Element, shape_id: int
) -> etree._Element | None:
    """Recursively find shape element by ID."""
    for shape_el in findall_v(parent, "Shape"):
        id_str = shape_el.get("ID")
        if id_str and int(id_str) == shape_id:
            return shape_el
        # Recurse into groups
        for shapes_container in findall_v(shape_el, "Shapes"):
            found = _find_shape_recursive(shapes_container, shape_id)
            if found is not None:
                return found
    return None


def get_shape_data(
    pkg: VisioPackage, page_num: int, shape_id: int
) -> list[ShapeDataProperty]:
    """Get custom properties (Property section) for a shape.

    Returns list of ShapeDataProperty with label, value, prompt, type, format.
    """
    shape_el = find_shape_element(pkg, page_num, shape_id)
    if shape_el is None:
        raise ValueError(f"Shape {shape_id} not found on page {page_num}")

    rows = get_section_rows(shape_el, "Property")
    results = []
    for row in rows:
        cells = row["cells"]
        results.append(
            ShapeDataProperty(
                label=cells.get("Label", {}).get("value"),
                value=cells.get("Value", {}).get("value"),
                prompt=cells.get("Prompt", {}).get("value"),
                type=cells.get("Type", {}).get("value"),
                format=cells.get("Format", {}).get("value"),
                sort_key=cells.get("SortKey", {}).get("value"),
                row_name=row["name"],
            )
        )
    return results


def get_shape_cells(
    pkg: VisioPackage, page_num: int, shape_id: int
) -> list[ShapeCellInfo]:
    """Get all singleton cells for a shape (full ShapeSheet dump).

    Returns list of ShapeCellInfo with name, value, formula, unit.
    """
    shape_el = find_shape_element(pkg, page_num, shape_id)
    if shape_el is None:
        raise ValueError(f"Shape {shape_id} not found on page {page_num}")

    raw_cells = get_all_cells(shape_el)
    return [
        ShapeCellInfo(
            name=c["name"],
            value=c["value"],
            formula=c["formula"],
            unit=c["unit"],
        )
        for c in raw_cells
    ]
