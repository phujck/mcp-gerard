"""Speaker notes operations for PowerPoint."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import CT, NSMAP, RT, qn
from mcp_handley_lab.microsoft.powerpoint.models import NotesInfo
from mcp_handley_lab.microsoft.powerpoint.ops.text import extract_text_from_txBody
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage


def get_notes(pkg: PowerPointPackage, slide_num: int) -> NotesInfo | None:
    """Get speaker notes for a slide."""
    notes_xml = pkg.get_notes_xml(slide_num)
    if notes_xml is None:
        return None

    # Find body placeholder in notes slide
    sp_tree = notes_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        return None

    body_text = ""
    paragraph_count = 0

    for sp in sp_tree.findall(qn("p:sp"), NSMAP):
        nvSpPr = sp.find(qn("p:nvSpPr"), NSMAP)
        if nvSpPr is None:
            continue

        nvPr = nvSpPr.find(qn("p:nvPr"), NSMAP)
        if nvPr is None:
            continue

        ph = nvPr.find(qn("p:ph"), NSMAP)
        if ph is None:
            continue

        # Body placeholder in notes slide
        ph_type = ph.get("type")
        if ph_type == "body":
            txBody = sp.find(qn("p:txBody"), NSMAP)
            if txBody is not None:
                body_text = extract_text_from_txBody(txBody)
                paragraph_count = len(txBody.findall(qn("a:p"), NSMAP))
            break

    if not body_text:
        return None

    return NotesInfo(
        slide_number=slide_num,
        text=body_text,
        paragraph_count=paragraph_count,
    )


def set_notes(pkg: PowerPointPackage, slide_num: int, text: str) -> None:
    """Set speaker notes for a slide.

    Creates notesSlide if it doesn't exist with proper bidirectional relationships.
    """
    notes_xml = pkg.get_notes_xml(slide_num)

    if notes_xml is None:
        # Create new notes slide
        notes_xml = _create_notes_slide(pkg, slide_num)

    # Find body placeholder
    sp_tree = notes_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
    if sp_tree is None:
        raise ValueError("Invalid notes slide structure")

    body_sp = None
    for sp in sp_tree.findall(qn("p:sp"), NSMAP):
        nvSpPr = sp.find(qn("p:nvSpPr"), NSMAP)
        if nvSpPr is None:
            continue

        nvPr = nvSpPr.find(qn("p:nvPr"), NSMAP)
        if nvPr is None:
            continue

        ph = nvPr.find(qn("p:ph"), NSMAP)
        if ph is None:
            continue

        if ph.get("type") == "body":
            body_sp = sp
            break

    if body_sp is None:
        # Create body placeholder if missing
        body_sp = _create_body_placeholder(sp_tree)

    # Update text
    _set_txBody_text(body_sp, text)

    # Mark dirty
    slide_partname = pkg.get_slide_partname(slide_num)
    slide_rels = pkg.get_rels(slide_partname)
    notes_rid = slide_rels.rId_for_reltype(RT.NOTES_SLIDE)
    if notes_rid:
        notes_path = pkg.resolve_rel_target(slide_partname, notes_rid)
        pkg.mark_xml_dirty(notes_path)


def _create_notes_slide(pkg: PowerPointPackage, slide_num: int) -> etree._Element:
    """Create a new notes slide with proper relationships."""
    slide_partname = pkg.get_slide_partname(slide_num)

    # Determine notes slide path (avoid collisions after deletions/reorders)
    notes_path = pkg.next_partname("/ppt/notesSlides/notesSlide", ".xml")

    # Create notes slide XML
    notes = etree.Element(
        qn("p:notes"),
        nsmap={
            "p": NSMAP["p"],
            "a": NSMAP["a"],
            "r": NSMAP["r"],
        },
    )

    cSld = etree.SubElement(notes, qn("p:cSld"))
    spTree = etree.SubElement(cSld, qn("p:spTree"))

    # Required group properties
    nvGrpSpPr = etree.SubElement(spTree, qn("p:nvGrpSpPr"))
    etree.SubElement(nvGrpSpPr, qn("p:cNvPr"), id="1", name="")
    etree.SubElement(nvGrpSpPr, qn("p:cNvGrpSpPr"))
    etree.SubElement(nvGrpSpPr, qn("p:nvPr"))
    etree.SubElement(spTree, qn("p:grpSpPr"))

    # Add slide image placeholder (optional but expected)
    _create_slide_image_placeholder(spTree)

    # Add body placeholder for notes text
    _create_body_placeholder(spTree)

    pkg.set_xml(notes_path, notes)

    # Add content type
    pkg._content_types[notes_path] = CT.PML_NOTES_SLIDE

    # Create slide → notesSlide relationship
    slide_rels = pkg.get_rels(slide_partname)
    slide_rels.get_or_add(RT.NOTES_SLIDE, notes_path)

    # Create notesSlide → slide relationship (bidirectional)
    notes_rels = pkg.get_rels(notes_path)
    notes_rels.get_or_add(RT.SLIDE, slide_partname)

    # Find and add notesMaster relationship
    pres_rels = pkg.get_rels(pkg.presentation_path)
    notes_master_rid = pres_rels.rId_for_reltype(RT.NOTES_MASTER)
    if notes_master_rid:
        notes_master_path = pkg.resolve_rel_target(
            pkg.presentation_path, notes_master_rid
        )
        notes_rels.get_or_add(RT.NOTES_MASTER, notes_master_path)

    return notes


def _create_slide_image_placeholder(spTree: etree._Element) -> etree._Element:
    """Create slide image placeholder in notes."""
    sp = etree.SubElement(spTree, qn("p:sp"))

    # nvSpPr
    nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
    etree.SubElement(nvSpPr, qn("p:cNvPr"), id="2", name="Slide Image")
    etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
    nvPr = etree.SubElement(nvSpPr, qn("p:nvPr"))
    etree.SubElement(nvPr, qn("p:ph"), type="sldImg")

    # spPr
    etree.SubElement(sp, qn("p:spPr"))

    return sp


def _create_body_placeholder(spTree: etree._Element) -> etree._Element:
    """Create body placeholder for notes text."""
    sp = etree.SubElement(spTree, qn("p:sp"))

    # nvSpPr
    nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
    etree.SubElement(nvSpPr, qn("p:cNvPr"), id="3", name="Notes Placeholder")
    etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
    nvPr = etree.SubElement(nvSpPr, qn("p:nvPr"))
    etree.SubElement(nvPr, qn("p:ph"), type="body", idx="1")

    # spPr
    etree.SubElement(sp, qn("p:spPr"))

    # txBody with empty paragraph
    txBody = etree.SubElement(sp, qn("p:txBody"))
    etree.SubElement(txBody, qn("a:bodyPr"))
    etree.SubElement(txBody, qn("a:lstStyle"))
    p = etree.SubElement(txBody, qn("a:p"))
    etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")

    return sp


def _set_txBody_text(sp: etree._Element, text: str) -> None:
    """Set text in a shape's txBody, preserving formatting where possible."""
    txBody = sp.find(qn("p:txBody"), NSMAP)

    if txBody is None:
        # Create txBody
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:lstStyle"))

    # Get existing paragraph properties if any
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

    # Add new paragraphs from text
    paragraphs = text.split("\n")
    for para_text in paragraphs:
        p = etree.SubElement(txBody, qn("a:p"))

        # Restore paragraph properties if we had them
        if existing_pPr is not None:
            p.append(existing_pPr.__copy__())

        if para_text:
            r = etree.SubElement(p, qn("a:r"))

            # Restore run properties if we had them
            if existing_rPr is not None:
                r.append(existing_rPr.__copy__())
            else:
                etree.SubElement(r, qn("a:rPr"), lang="en-US")

            t = etree.SubElement(r, qn("a:t"))
            t.text = para_text
        else:
            etree.SubElement(p, qn("a:endParaRPr"), lang="en-US")
