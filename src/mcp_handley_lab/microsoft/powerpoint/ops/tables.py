"""Table operations for PowerPoint."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import NSMAP, qn
from mcp_handley_lab.microsoft.powerpoint.ops.core import (
    get_shape_id,
    inches_to_emu,
    make_shape_key,
    parse_shape_key,
)
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage


def add_table(
    pkg: PowerPointPackage,
    slide_num: int,
    rows: int,
    cols: int,
    x: float = 1.0,
    y: float = 1.0,
    width: float = 6.0,
    height: float = 2.0,
) -> str:
    """Add a table to a slide.

    Args:
        pkg: PowerPoint package
        slide_num: Target slide number
        rows: Number of rows
        cols: Number of columns
        x: X position in inches
        y: Y position in inches
        width: Table width in inches
        height: Table height in inches

    Returns:
        Shape key for the new table (slide_num:shape_id)
    """
    if rows < 1 or cols < 1:
        raise ValueError("Table must have at least 1 row and 1 column")

    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)
    sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        raise ValueError(f"Slide {slide_num} has no spTree")

    # Get next shape ID
    max_id = 0
    for sp in sp_tree.findall(".//" + qn("p:cNvPr"), NSMAP):
        id_str = sp.get("id", "0")
        if id_str.isdigit():
            max_id = max(max_id, int(id_str))
    new_id = max_id + 1

    # Calculate cell dimensions
    cell_width = inches_to_emu(width) // cols
    cell_height = inches_to_emu(height) // rows

    # Create graphicFrame element
    graphic_frame = _create_table_graphic_frame(
        new_id, rows, cols, x, y, width, height, cell_width, cell_height
    )
    sp_tree.append(graphic_frame)

    pkg.mark_xml_dirty(slide_partname)
    return make_shape_key(slide_num, new_id)


def set_table_cell(
    pkg: PowerPointPackage,
    shape_key: str,
    row: int,
    col: int,
    text: str,
) -> bool:
    """Set text in a table cell.

    Args:
        pkg: PowerPoint package
        shape_key: Table shape key (slide_num:shape_id)
        row: Row index (0-based)
        col: Column index (0-based)
        text: Text to set

    Returns:
        True if successful, False if cell not found
    """
    slide_num, shape_id = parse_shape_key(shape_key)
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)
    sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        return False

    # Find the graphic frame with matching shape_id
    for gf in sp_tree.findall(qn("p:graphicFrame"), NSMAP):
        gf_id = get_shape_id(gf)
        if gf_id == shape_id:
            # Find the table
            tbl = gf.find(
                qn("a:graphic") + "/" + qn("a:graphicData") + "/" + qn("a:tbl"), NSMAP
            )
            if tbl is None:
                return False

            # Find the cell
            tr_list = tbl.findall(qn("a:tr"), NSMAP)
            if row >= len(tr_list):
                return False

            tc_list = tr_list[row].findall(qn("a:tc"), NSMAP)
            if col >= len(tc_list):
                return False

            tc = tc_list[col]

            # Update the cell text
            _set_cell_text(tc, text)
            pkg.mark_xml_dirty(slide_partname)
            return True

    return False


def _create_table_graphic_frame(
    shape_id: int,
    rows: int,
    cols: int,
    x: float,
    y: float,
    width: float,
    height: float,
    cell_width: int,
    cell_height: int,
) -> etree._Element:
    """Create a graphicFrame containing a table."""
    gf = etree.Element(qn("p:graphicFrame"), nsmap={"p": NSMAP["p"], "a": NSMAP["a"]})

    # nvGraphicFramePr
    nvGfPr = etree.SubElement(gf, qn("p:nvGraphicFramePr"))
    cNvPr = etree.SubElement(nvGfPr, qn("p:cNvPr"))
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", f"Table {shape_id - 1}")
    cNvGfPr = etree.SubElement(nvGfPr, qn("p:cNvGraphicFramePr"))
    etree.SubElement(cNvGfPr, qn("a:graphicFrameLocks"), noGrp="1")
    etree.SubElement(nvGfPr, qn("p:nvPr"))

    # xfrm (position and size)
    xfrm = etree.SubElement(gf, qn("p:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(inches_to_emu(x)))
    off.set("y", str(inches_to_emu(y)))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(inches_to_emu(width)))
    ext.set("cy", str(inches_to_emu(height)))

    # graphic container
    graphic = etree.SubElement(gf, qn("a:graphic"))
    graphicData = etree.SubElement(graphic, qn("a:graphicData"))
    graphicData.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/table")

    # Table element
    tbl = etree.SubElement(graphicData, qn("a:tbl"))

    # Table properties
    tblPr = etree.SubElement(tbl, qn("a:tblPr"))
    tblPr.set("firstRow", "1")
    tblPr.set("bandRow", "1")

    # Table grid (column widths)
    tblGrid = etree.SubElement(tbl, qn("a:tblGrid"))
    for _ in range(cols):
        gridCol = etree.SubElement(tblGrid, qn("a:gridCol"))
        gridCol.set("w", str(cell_width))

    # Table rows
    for _ in range(rows):
        tr = etree.SubElement(tbl, qn("a:tr"))
        tr.set("h", str(cell_height))
        for _ in range(cols):
            tc = _create_empty_cell()
            tr.append(tc)

    return gf


def _create_empty_cell() -> etree._Element:
    """Create an empty table cell."""
    tc = etree.Element(qn("a:tc"))

    # txBody with empty paragraph
    txBody = etree.SubElement(tc, qn("a:txBody"))
    etree.SubElement(txBody, qn("a:bodyPr"))
    etree.SubElement(txBody, qn("a:lstStyle"))
    p = etree.SubElement(txBody, qn("a:p"))
    etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")

    # Cell properties
    etree.SubElement(tc, qn("a:tcPr"))

    return tc


def _set_cell_text(tc: etree._Element, text: str) -> None:
    """Set text in a table cell."""
    txBody = tc.find(qn("a:txBody"), NSMAP)
    if txBody is None:
        txBody = etree.SubElement(tc, qn("a:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:lstStyle"))

    # Remove existing paragraphs
    for p in list(txBody.findall(qn("a:p"), NSMAP)):
        txBody.remove(p)

    # Add new paragraphs
    for line in text.split("\n"):
        p = etree.SubElement(txBody, qn("a:p"))
        if line:
            r = etree.SubElement(p, qn("a:r"))
            etree.SubElement(r, qn("a:rPr"), lang="en-US")
            t = etree.SubElement(r, qn("a:t"))
            t.text = line
        else:
            etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")
