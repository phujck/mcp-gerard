"""Text extraction operations for PowerPoint."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import NSMAP, qn

if TYPE_CHECKING:
    from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage


def extract_text_from_txBody(txBody: etree._Element | None) -> str:
    """Extract text content from a txBody element.

    Handles:
    - a:p paragraphs (separated by newlines)
    - a:r runs with a:t text
    - a:fld fields (date, slide number, etc.)
    - a:br line breaks
    """
    if txBody is None:
        return ""

    paragraphs = []

    for p in txBody.findall(qn("a:p"), NSMAP):
        parts = []

        for child in p:
            tag = etree.QName(child).localname

            if tag == "r" or tag == "fld":  # Run or field
                t = child.find(qn("a:t"), NSMAP)
                if t is not None and t.text:
                    parts.append(t.text)

            elif tag == "br":  # Line break
                parts.append("\n")

            elif tag == "tab":  # Tab character
                parts.append("\t")

        if parts:
            paragraphs.append("".join(parts))

    return "\n".join(paragraphs)


def extract_text_from_table(tbl: etree._Element) -> str:
    """Extract text from a table (a:tbl).

    Returns rows separated by newlines, cells separated by tabs.
    """
    rows = []

    for tr in tbl.findall(qn("a:tr"), NSMAP):
        cells = []
        for tc in tr.findall(qn("a:tc"), NSMAP):
            txBody = tc.find(qn("a:txBody"), NSMAP)
            cell_text = extract_text_from_txBody(txBody)
            # Replace internal newlines with spaces in cells
            cells.append(cell_text.replace("\n", " "))
        rows.append("\t".join(cells))

    return "\n".join(rows)


def extract_text_from_shape(shape: etree._Element) -> str:
    """Extract text from any shape element."""
    tag = etree.QName(shape).localname

    if tag == "sp":  # Shape
        txBody = shape.find(qn("p:txBody"), NSMAP)
        return extract_text_from_txBody(txBody)

    elif tag == "graphicFrame":  # Could be table, chart, etc.
        # Check for table
        tbl = shape.find(".//" + qn("a:tbl"), NSMAP)
        if tbl is not None:
            return extract_text_from_table(tbl)
        # Charts don't have extractable text
        return ""

    elif tag == "grpSp":  # Group
        # Extract text from all children
        texts = []
        sp_tree_children = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"]
        for child in shape:
            child_tag = etree.QName(child).localname
            if child_tag in sp_tree_children:
                text = extract_text_from_shape(child)
                if text:
                    texts.append(text)
        return "\n".join(texts)

    return ""


def extract_title_text(slide_xml: etree._Element) -> str | None:
    """Extract title text from slide's title placeholder."""
    sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        return None

    for sp in sp_tree.findall(qn("p:sp"), NSMAP):
        # Check for title placeholder
        nvSpPr = sp.find(qn("p:nvSpPr"), NSMAP)
        if nvSpPr is None:
            continue

        nvPr = nvSpPr.find(qn("p:nvPr"), NSMAP)
        if nvPr is None:
            continue

        ph = nvPr.find(qn("p:ph"), NSMAP)
        if ph is None:
            continue

        ph_type = ph.get("type")
        if ph_type in ("title", "ctrTitle"):
            txBody = sp.find(qn("p:txBody"), NSMAP)
            text = extract_text_from_txBody(txBody)
            if text:
                return text.strip()

    return None


def extract_all_text_from_slide(slide_xml: etree._Element) -> str:
    """Extract all text from a slide in reading order.

    Note: This does NOT do spatial sorting - use shapes module for that.
    """
    sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        return ""

    texts = []
    shape_tags = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"]

    for child in sp_tree:
        child_tag = etree.QName(child).localname
        if child_tag in shape_tags:
            text = extract_text_from_shape(child)
            if text:
                texts.append(text)

    return "\n\n".join(texts)


def set_shape_text(
    sp: etree._Element,
    text: str,
    bullet_style: str | None = None,
) -> None:
    """Set text in a shape, preserving formatting.

    Works with any p:sp element. Creates txBody if not present.
    Preserves paragraph and run formatting from existing content.
    Supports tab characters (\\t) via a:tab elements.

    Args:
        sp: Shape element (p:sp)
        text: Text content (newlines separate paragraphs, tabs create a:tab elements)
        bullet_style: Optional bullet style for all paragraphs:
            - "bullet": Standard bullet character (U+2022)
            - "dash": Dash character (U+2013)
            - "number": Auto-numbered (arabicPeriod)
            - "none": Explicitly remove bullets
            - None: Don't modify bullet state
    """
    import copy

    txBody = sp.find(qn("p:txBody"), NSMAP)

    if txBody is None:
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:lstStyle"))

    # Preserve first paragraph/run/endPara properties
    existing_p = txBody.find(qn("a:p"), NSMAP)
    existing_pPr = None
    existing_rPr = None
    existing_endParaRPr = None

    if existing_p is not None:
        pPr = existing_p.find(qn("a:pPr"), NSMAP)
        if pPr is not None:
            existing_pPr = copy.deepcopy(pPr)
        endParaRPr = existing_p.find(qn("a:endParaRPr"), NSMAP)
        if endParaRPr is not None:
            existing_endParaRPr = copy.deepcopy(endParaRPr)
        existing_r = existing_p.find(qn("a:r"), NSMAP)
        if existing_r is not None:
            rPr = existing_r.find(qn("a:rPr"), NSMAP)
            if rPr is not None:
                existing_rPr = copy.deepcopy(rPr)

    # Remove existing paragraphs
    for p in list(txBody.findall(qn("a:p"), NSMAP)):
        txBody.remove(p)

    # Add new paragraphs
    paragraphs = text.split("\n")
    for para_text in paragraphs:
        p = etree.SubElement(txBody, qn("a:p"))

        if existing_pPr is not None:
            pPr = copy.deepcopy(existing_pPr)
            p.append(pPr)
        else:
            pPr = None

        # Apply bullet styling if requested
        if bullet_style is not None:
            if pPr is None:
                pPr = etree.SubElement(p, qn("a:pPr"))
            _apply_bullet_style(pPr, bullet_style)

        if para_text:
            # Split on tabs and create a:tab elements between segments
            segments = para_text.split("\t")
            for i, segment in enumerate(segments):
                if i > 0:
                    etree.SubElement(p, qn("a:tab"))
                if segment:
                    r = etree.SubElement(p, qn("a:r"))
                    if existing_rPr is not None:
                        r.append(copy.deepcopy(existing_rPr))
                    else:
                        etree.SubElement(r, qn("a:rPr"), lang="en-US")
                    t = etree.SubElement(r, qn("a:t"))
                    t.text = segment

        # Always append a:endParaRPr (PowerPoint expects it on every paragraph)
        if existing_endParaRPr is not None:
            p.append(copy.deepcopy(existing_endParaRPr))
        else:
            etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")


# Bullet element tags that must be removed before applying new bullet style
_BULLET_TAGS = (
    "a:buChar",
    "a:buAutoNum",
    "a:buNone",
    "a:buFont",
    "a:buBlip",
    "a:buSzPts",
    "a:buSzPct",
    "a:buClrTx",
    "a:buClr",
)

# EMUs for bullet indentation: 0.25 inches
_BULLET_INDENT_EMU = 228600


_VALID_BULLET_STYLES = {"bullet", "dash", "number", "none"}


def _apply_bullet_style(pPr: etree._Element, style: str) -> None:
    """Apply bullet style to a paragraph properties element.

    Merges bullet properties into existing pPr, preserving other properties
    like alignment and spacing.
    """
    if style not in _VALID_BULLET_STYLES:
        raise ValueError(
            f"Unknown bullet_style '{style}'. "
            f"Valid values: {', '.join(sorted(_VALID_BULLET_STYLES))}"
        )

    # Remove all existing bullet elements
    for tag in _BULLET_TAGS:
        existing = pPr.find(qn(tag), NSMAP)
        if existing is not None:
            pPr.remove(existing)

    if style == "none":
        etree.SubElement(pPr, qn("a:buNone"))
        # Remove indent/margin if explicitly removing bullets
        pPr.attrib.pop("marL", None)
        pPr.attrib.pop("indent", None)
    elif style == "bullet":
        pPr.set("marL", str(_BULLET_INDENT_EMU))
        pPr.set("indent", str(-_BULLET_INDENT_EMU))
        etree.SubElement(pPr, qn("a:buChar"), char="\u2022")
    elif style == "dash":
        pPr.set("marL", str(_BULLET_INDENT_EMU))
        pPr.set("indent", str(-_BULLET_INDENT_EMU))
        etree.SubElement(pPr, qn("a:buChar"), char="\u2013")
    elif style == "number":
        pPr.set("marL", str(_BULLET_INDENT_EMU))
        pPr.set("indent", str(-_BULLET_INDENT_EMU))
        etree.SubElement(pPr, qn("a:buAutoNum"), type="arabicPeriod")


def add_hyperlink(
    pkg: PowerPointPackage,
    shape_key: str,
    url: str,
    tooltip: str | None = None,
) -> bool:
    """Add a hyperlink to all text runs in a shape.

    Args:
        pkg: PowerPoint package
        shape_key: Shape identifier (slide_num:shape_id)
        url: URL for the hyperlink
        tooltip: Optional tooltip text

    Returns:
        True if hyperlink was added, False if shape not found or has no text runs
    """
    from mcp_handley_lab.microsoft.powerpoint.constants import RT
    from mcp_handley_lab.microsoft.powerpoint.ops.core import (
        find_shape_by_id,
        parse_shape_key,
    )

    slide_num, shape_id = parse_shape_key(shape_key)
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_xml = pkg.get_slide_xml(slide_num)

    shape = find_shape_by_id(slide_xml, shape_id)
    if shape is None:
        return False

    txBody = shape.find(qn("p:txBody"), NSMAP)
    if txBody is None:
        return False

    # Find all runs (a:r) and fields (a:fld) across all paragraphs
    elements = txBody.findall(".//" + qn("a:r"), NSMAP) + txBody.findall(
        ".//" + qn("a:fld"), NSMAP
    )
    if not elements:
        return False

    # Add external hyperlink relationship to slide
    slide_rels = pkg.get_rels(slide_partname)
    rId = slide_rels.get_or_add(RT.HYPERLINK, url, is_external=True)

    # Apply hyperlink to each run/field
    for elem in elements:
        # Get or create rPr
        rPr = elem.find(qn("a:rPr"), NSMAP)
        if rPr is None:
            rPr = etree.Element(qn("a:rPr"), lang="en-US")
            elem.insert(0, rPr)

        # Remove existing hyperlink
        existing_hlink = rPr.find(qn("a:hlinkClick"), NSMAP)
        if existing_hlink is not None:
            rPr.remove(existing_hlink)

        # Add new hyperlink
        hlink = etree.SubElement(rPr, qn("a:hlinkClick"))
        hlink.set(qn("r:id"), rId)
        if tooltip:
            hlink.set("tooltip", tooltip)

    pkg.mark_xml_dirty(slide_partname)
    pkg._dirty_rels.add(slide_partname)
    return True
