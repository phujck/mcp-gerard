"""Shape styling operations for PowerPoint."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import NSMAP, qn
from mcp_handley_lab.microsoft.powerpoint.ops.core import (
    find_shape_by_id,
    parse_shape_key,
)
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage

# EMUs per point for line width
EMU_PER_PT = 12700


def set_slide_background(
    pkg: PowerPointPackage,
    slide_num: int,
    color: str,
) -> bool:
    """Set a solid color background on a slide.

    Args:
        pkg: PowerPoint package
        slide_num: Slide number (1-based)
        color: Hex color without # (e.g., "FF0000" for red)

    Returns:
        True if successful
    """
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)

    cSld = slide_xml.find(qn("p:cSld"), NSMAP)
    if cSld is None:
        return False

    # Remove existing background if present
    existing_bg = cSld.find(qn("p:bg"), NSMAP)
    if existing_bg is not None:
        cSld.remove(existing_bg)

    # Create p:bg element with explicit namespace bindings
    bg = etree.Element(
        qn("p:bg"), nsmap={"p": NSMAP["p"], "a": NSMAP["a"], "r": NSMAP["r"]}
    )
    bgPr = etree.SubElement(bg, qn("p:bgPr"))
    solid_fill = etree.SubElement(bgPr, qn("a:solidFill"))
    srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
    srgb_clr.set("val", color.upper().lstrip("#"))
    etree.SubElement(bgPr, qn("a:effectLst"))

    # Insert as first child of p:cSld (before p:spTree)
    cSld.insert(0, bg)

    pkg.mark_xml_dirty(slide_partname)
    return True


def set_shape_fill(
    pkg: PowerPointPackage,
    shape_key: str,
    color: str,
) -> bool:
    """Set the fill color of a shape.

    Args:
        pkg: PowerPoint package
        shape_key: Shape identifier (slide_num:shape_id)
        color: Hex color without # (e.g., "FF0000" for red)

    Returns:
        True if successful, False if shape not found
    """
    slide_num, shape_id = parse_shape_key(shape_key)
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)

    shape = find_shape_by_id(slide_xml, shape_id)
    if shape is None:
        return False

    # Only apply fill to standard shapes and connectors
    tag = etree.QName(shape.tag).localname
    if tag not in ("sp", "cxnSp"):
        return False

    # Get or create spPr
    spPr = shape.find(qn("p:spPr"), NSMAP)
    if spPr is None:
        spPr = etree.SubElement(shape, qn("p:spPr"))

    # Remove existing fill elements
    for fill_tag in (
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:pattFill",
    ):
        existing = spPr.find(qn(fill_tag), NSMAP)
        if existing is not None:
            spPr.remove(existing)

    # Add solidFill with color
    solid_fill = etree.SubElement(spPr, qn("a:solidFill"))
    srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
    srgb_clr.set("val", color.upper().lstrip("#"))

    pkg.mark_xml_dirty(slide_partname)
    return True


def set_shape_line(
    pkg: PowerPointPackage,
    shape_key: str,
    color: str | None = None,
    width: float | None = None,
) -> bool:
    """Set the outline/border of a shape.

    Args:
        pkg: PowerPoint package
        shape_key: Shape identifier (slide_num:shape_id)
        color: Hex color without # (e.g., "000000" for black)
        width: Line width in points

    Returns:
        True if successful, False if shape not found
    """
    slide_num, shape_id = parse_shape_key(shape_key)
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)

    shape = find_shape_by_id(slide_xml, shape_id)
    if shape is None:
        return False

    # Get or create spPr
    spPr = shape.find(qn("p:spPr"), NSMAP)
    if spPr is None:
        spPr = etree.SubElement(shape, qn("p:spPr"))

    # Get or create ln element
    ln = spPr.find(qn("a:ln"), NSMAP)
    if ln is None:
        ln = etree.SubElement(spPr, qn("a:ln"))

    # Set width if provided
    if width is not None:
        ln.set("w", str(int(width * EMU_PER_PT)))

    # Set color if provided
    if color is not None:
        # Remove existing fill in line
        for fill_tag in ("a:noFill", "a:solidFill", "a:gradFill"):
            existing = ln.find(qn(fill_tag), NSMAP)
            if existing is not None:
                ln.remove(existing)

        # Add solidFill with color
        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
        srgb_clr.set("val", color.upper().lstrip("#"))

    pkg.mark_xml_dirty(slide_partname)
    return True


def set_text_style(
    pkg: PowerPointPackage,
    shape_key: str,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    alignment: str | None = None,
) -> bool:
    """Set text style properties for all text in a shape.

    Args:
        pkg: PowerPoint package
        shape_key: Shape identifier (slide_num:shape_id)
        size: Font size in points
        bold: Bold text
        italic: Italic text
        color: Hex color without # (e.g., "000000" for black)
        alignment: Text alignment ("left", "center", "right", "justify")

    Returns:
        True if successful, False if shape not found or has no text
    """
    slide_num, shape_id = parse_shape_key(shape_key)
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)

    shape = find_shape_by_id(slide_xml, shape_id)
    if shape is None:
        return False

    # Find txBody
    txBody = shape.find(qn("p:txBody"), NSMAP)
    if txBody is None:
        return False

    # Map alignment values to OOXML values
    align_map = {
        "left": "l",
        "center": "ctr",
        "right": "r",
        "justify": "just",
    }

    # Apply styles to all paragraphs and runs
    paragraphs = txBody.findall(qn("a:p"), NSMAP)
    if not paragraphs:
        return False

    for p in paragraphs:
        # Apply alignment at paragraph level
        if alignment is not None:
            pPr = p.find(qn("a:pPr"), NSMAP)
            if pPr is None:
                # Insert pPr as first child
                pPr = etree.Element(qn("a:pPr"))
                p.insert(0, pPr)
            algn_val = align_map.get(alignment.lower(), "l")
            pPr.set("algn", algn_val)

        # Apply font styles to all runs
        runs = p.findall(qn("a:r"), NSMAP)
        for r in runs:
            # Get or create rPr
            rPr = r.find(qn("a:rPr"), NSMAP)
            if rPr is None:
                # Insert rPr before text element
                rPr = etree.Element(qn("a:rPr"))
                r.insert(0, rPr)

            # Set font size (in 100ths of a point)
            if size is not None:
                rPr.set("sz", str(int(size * 100)))

            # Set bold
            if bold is not None:
                rPr.set("b", "1" if bold else "0")

            # Set italic
            if italic is not None:
                rPr.set("i", "1" if italic else "0")

            # Set color
            if color is not None:
                # Remove existing fill in rPr
                for fill_tag in ("a:noFill", "a:solidFill", "a:gradFill"):
                    existing = rPr.find(qn(fill_tag), NSMAP)
                    if existing is not None:
                        rPr.remove(existing)

                # Add solidFill with color
                solid_fill = etree.SubElement(rPr, qn("a:solidFill"))
                srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb_clr.set("val", color.upper().lstrip("#"))

        # Also apply to endParaRPr if it exists (for empty paragraphs)
        endParaRPr = p.find(qn("a:endParaRPr"), NSMAP)
        if endParaRPr is not None:
            if size is not None:
                endParaRPr.set("sz", str(int(size * 100)))
            if bold is not None:
                endParaRPr.set("b", "1" if bold else "0")
            if italic is not None:
                endParaRPr.set("i", "1" if italic else "0")

    pkg.mark_xml_dirty(slide_partname)
    return True
