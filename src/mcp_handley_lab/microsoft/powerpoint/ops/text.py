"""Text extraction operations for PowerPoint."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import NSMAP, qn


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

            if tag == "r" or tag == "fld":  # Run
                t = child.find(qn("a:t"), NSMAP)
                if t is not None and t.text:
                    parts.append(t.text)

            elif tag == "br":  # Line break
                parts.append("\n")

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


def set_shape_text(sp: etree._Element, text: str) -> None:
    """Set text in a shape, preserving formatting.

    Works with any p:sp element. Creates txBody if not present.
    Preserves paragraph and run formatting from existing content.
    """
    txBody = sp.find(qn("p:txBody"), NSMAP)

    if txBody is None:
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:lstStyle"))

    # Preserve first paragraph/run properties
    existing_p = txBody.find(qn("a:p"), NSMAP)
    existing_pPr = None
    existing_rPr = None

    if existing_p is not None:
        existing_pPr = existing_p.find(qn("a:pPr"), NSMAP)
        existing_r = existing_p.find(qn("a:r"), NSMAP)
        if existing_r is not None:
            existing_rPr = existing_r.find(qn("a:rPr"), NSMAP)

    # Remove existing paragraphs
    for p in list(txBody.findall(qn("a:p"), NSMAP)):
        txBody.remove(p)

    # Add new paragraphs
    paragraphs = text.split("\n")
    for para_text in paragraphs:
        p = etree.SubElement(txBody, qn("a:p"))

        if existing_pPr is not None:
            p.append(existing_pPr.__copy__())

        if para_text:
            r = etree.SubElement(p, qn("a:r"))
            if existing_rPr is not None:
                r.append(existing_rPr.__copy__())
            else:
                etree.SubElement(r, qn("a:rPr"), lang="en-US")
            t = etree.SubElement(r, qn("a:t"))
            t.text = para_text
        else:
            etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")
