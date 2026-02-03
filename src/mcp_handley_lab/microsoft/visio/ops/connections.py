"""Connector relationship parsing for Visio."""

from __future__ import annotations

from mcp_handley_lab.microsoft.visio.constants import findall_v
from mcp_handley_lab.microsoft.visio.models import ConnectionInfo
from mcp_handley_lab.microsoft.visio.ops.core import extract_shape_text
from mcp_handley_lab.microsoft.visio.package import VisioPackage


def list_connections(pkg: VisioPackage, page_num: int) -> list[ConnectionInfo]:
    """List connector relationships on a page.

    Parses <Connect> elements from the page XML. These are direct children of
    <PageContents> and define which shapes are connected via connectors.

    Each connector has two Connect elements: one for BeginX (start) and one
    for EndX (end). We group by FromSheet (connector shape ID) to build
    complete connection records.

    Limitations: Some advanced connection types (dynamic glue, inherited
    connections from masters) may not be fully captured.
    """
    page_xml = pkg.get_page_xml(page_num)

    # Build shape ID -> name mapping for the page (recursive for groups)
    shape_names: dict[int, str] = {}
    shape_texts: dict[int, str | None] = {}

    def _index_shapes(parent) -> None:
        for shape_el in findall_v(parent, "Shape"):
            id_str = shape_el.get("ID")
            if id_str is None:
                continue
            sid = int(id_str)
            shape_names[sid] = shape_el.get("NameU") or shape_el.get("Name") or ""
            shape_texts[sid] = extract_shape_text(shape_el)
            # Recurse into group children
            for child_container in findall_v(shape_el, "Shapes"):
                _index_shapes(child_container)

    for container in findall_v(page_xml, "Shapes"):
        _index_shapes(container)

    # Parse Connect elements
    # Group by FromSheet (connector ID)
    connector_ends: dict[int, dict[str, int | None]] = {}

    connects = findall_v(page_xml, "Connect")
    for connect in connects:
        from_sheet_str = connect.get("FromSheet")
        to_sheet_str = connect.get("ToSheet")
        from_cell = connect.get("FromCell")

        if from_sheet_str is None or to_sheet_str is None:
            continue

        from_sheet = int(from_sheet_str)
        to_sheet = int(to_sheet_str)

        if from_sheet not in connector_ends:
            connector_ends[from_sheet] = {"begin": None, "end": None}

        # FromCell indicates which end of the connector
        if from_cell == "BeginX":
            connector_ends[from_sheet]["begin"] = to_sheet
        elif from_cell == "EndX":
            connector_ends[from_sheet]["end"] = to_sheet
        else:
            # Fallback: if FromCell is missing or unknown, try FromPart
            # FromPart 9 = BeginX, 12 = EndX
            from_part = connect.get("FromPart")
            if from_part == "9":
                connector_ends[from_sheet]["begin"] = to_sheet
            elif from_part == "12":
                connector_ends[from_sheet]["end"] = to_sheet
            else:
                # Best effort: assign to first empty slot
                if connector_ends[from_sheet]["begin"] is None:
                    connector_ends[from_sheet]["begin"] = to_sheet
                elif connector_ends[from_sheet]["end"] is None:
                    connector_ends[from_sheet]["end"] = to_sheet

    # Build ConnectionInfo records
    results = []
    for connector_id, ends in connector_ends.items():
        from_id = ends.get("begin")
        to_id = ends.get("end")

        results.append(
            ConnectionInfo(
                connector_id=connector_id,
                connector_name=shape_names.get(connector_id),
                connector_text=shape_texts.get(connector_id),
                from_shape_id=from_id,
                from_shape_name=shape_names.get(from_id) if from_id else None,
                to_shape_id=to_id,
                to_shape_name=shape_names.get(to_id) if to_id else None,
            )
        )

    return results
