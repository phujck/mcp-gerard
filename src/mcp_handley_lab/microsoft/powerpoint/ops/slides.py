"""Slide operations for PowerPoint."""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.powerpoint.constants import NSMAP, RT, qn
from mcp_handley_lab.microsoft.powerpoint.models import LayoutInfo, SlideInfo
from mcp_handley_lab.microsoft.powerpoint.ops.text import extract_title_text
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage


def list_slides(pkg: PowerPointPackage) -> list[SlideInfo]:
    """List all slides with metadata."""
    slides = []

    for num, _rid, partname in pkg.get_slide_paths():
        slide_xml = pkg.get_xml(partname)

        # Count shapes in spTree
        sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
        shape_count = 0
        if sp_tree is not None:
            # Count direct shape children (excluding nvGrpSpPr which is the tree itself)
            for child in sp_tree:
                tag = etree.QName(child).localname
                if tag in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                    shape_count += 1

        # Check for notes
        has_notes = pkg.get_notes_xml(num) is not None

        # Get layout name
        layout_name = pkg.get_slide_layout_name(num)

        # Extract title from title placeholder
        title = extract_title_text(slide_xml)

        slides.append(
            SlideInfo(
                number=num,
                title=title,
                shape_count=shape_count,
                has_notes=has_notes,
                layout_name=layout_name,
            )
        )

    return slides


def get_slide_count(pkg: PowerPointPackage) -> int:
    """Get total number of slides."""
    return len(pkg.get_slide_paths())


def get_notes_count(pkg: PowerPointPackage) -> int:
    """Get number of slides with speaker notes."""
    count = 0
    for num, _rid, _partname in pkg.get_slide_paths():
        if pkg.get_notes_xml(num) is not None:
            count += 1
    return count


def list_layouts(pkg: PowerPointPackage) -> list[LayoutInfo]:
    """List all available slide layouts from all slide masters.

    Uses p:sldMasterIdLst for master ordering and p:sldLayoutIdLst for layout
    ordering within each master, as required by OOXML spec.
    """
    pres = pkg.presentation_xml
    layouts = []

    # Iterate over all masters in sldMasterIdLst (preserves ordering)
    sld_master_id_lst = pres.find(qn("p:sldMasterIdLst"), NSMAP)
    if sld_master_id_lst is None:
        return []

    for master_idx, master_id_el in enumerate(
        sld_master_id_lst.findall(qn("p:sldMasterId"), NSMAP)
    ):
        master_rid = master_id_el.get(qn("r:id"))
        if master_rid is None:
            continue

        master_path = pkg.resolve_rel_target(pkg.presentation_path, master_rid)
        master_xml = pkg.get_xml(master_path)

        # Get master name from cSld@name
        master_cSld = master_xml.find(qn("p:cSld"), NSMAP)
        master_name = (
            master_cSld.get("name", f"Master {master_idx + 1}")
            if master_cSld is not None
            else f"Master {master_idx + 1}"
        )

        master_rels = pkg.get_rels(master_path)

        # Use sldLayoutIdLst for proper ordering (not relationship iteration)
        sld_layout_id_lst = master_xml.find(qn("p:sldLayoutIdLst"), NSMAP)
        if sld_layout_id_lst is None:
            continue

        for layout_id_el in sld_layout_id_lst.findall(qn("p:sldLayoutId"), NSMAP):
            layout_rid = layout_id_el.get(qn("r:id"))
            if layout_rid is None or layout_rid not in master_rels:
                continue

            layout_path = pkg.resolve_rel_target(master_path, layout_rid)
            layout_xml = pkg.get_xml(layout_path)

            # Get layout name from cSld@name
            cSld = layout_xml.find(qn("p:cSld"), NSMAP)
            name = cSld.get("name", "Untitled") if cSld is not None else "Untitled"

            # Get layout type from root element attribute
            layout_type = layout_xml.get("type")

            # Collect placeholder info
            sp_tree = layout_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
            placeholder_count = 0
            placeholder_types = []
            if sp_tree is not None:
                for sp in sp_tree.findall(qn("p:sp"), NSMAP):
                    ph = sp.find(
                        qn("p:nvSpPr") + "/" + qn("p:nvPr") + "/" + qn("p:ph"), NSMAP
                    )
                    if ph is not None:
                        placeholder_count += 1
                        ph_type = ph.get("type", "body")
                        placeholder_types.append(ph_type)

            layouts.append(
                LayoutInfo(
                    name=name,
                    type=layout_type,
                    placeholder_count=placeholder_count,
                    placeholder_types=placeholder_types,
                    master_name=master_name,
                    master_index=master_idx,
                )
            )

    return layouts


def add_slide(
    pkg: PowerPointPackage,
    layout_name: str | None = None,
    position: int | None = None,
) -> int:
    """Add a new slide to the presentation.

    Args:
        pkg: PowerPoint package
        layout_name: Layout name to use (default: first layout)
        position: 1-based position to insert (default: end)

    Returns:
        New slide number
    """
    pres = pkg.presentation_xml
    pres_path = pkg.presentation_path
    pres_rels = pkg.get_rels(pres_path)

    # Find layout
    layout_path = _find_layout_by_name(pkg, layout_name)

    # Determine new slide path (avoid collisions after deletions)
    new_slide_path = pkg.next_partname("/ppt/slides/slide", ".xml")

    # Create minimal slide XML
    slide = _create_minimal_slide()
    pkg.set_xml(new_slide_path, slide)

    # Add content type
    from mcp_handley_lab.microsoft.powerpoint.constants import CT

    pkg._content_types[new_slide_path] = CT.PML_SLIDE

    # Add relationship from presentation to slide
    new_rid = pres_rels.get_or_add(RT.SLIDE, new_slide_path)

    # Add slide to sldIdLst
    sld_id_lst = pres.find(qn("p:sldIdLst"), NSMAP)
    if sld_id_lst is None:
        sld_id_lst = etree.SubElement(pres, qn("p:sldIdLst"))

    # Generate new slide ID (must be unique, >= 256)
    existing_ids = [int(el.get("id", "0")) for el in sld_id_lst]
    new_id = max(existing_ids, default=255) + 1
    if new_id < 256:
        new_id = 256

    # Create the sldId element
    sld_id = etree.Element(
        qn("p:sldId"),
        id=str(new_id),
        nsmap={"r": NSMAP["r"]},
    )
    sld_id.set(qn("r:id"), new_rid)

    # Insert at position or append to end
    existing_count = len(list(sld_id_lst))
    if position is not None:
        # Clamp position to valid range (1 to existing_count + 1)
        insert_idx = max(0, min(position - 1, existing_count))
        sld_id_lst.insert(insert_idx, sld_id)
    else:
        # Append to end
        sld_id_lst.append(sld_id)

    # Add relationship from slide to layout
    slide_rels = pkg.get_rels(new_slide_path)
    slide_rels.get_or_add(RT.SLIDE_LAYOUT, layout_path)

    # Mark dirty and invalidate caches
    pkg.mark_xml_dirty(pres_path)
    pkg.invalidate_caches()

    # Determine actual slide number after insertion
    # The slide's position is its index in sldIdLst + 1
    for idx, el in enumerate(sld_id_lst):
        if el.get(qn("r:id")) == new_rid:
            return idx + 1

    # Fallback (should not reach here)
    return existing_count + 1


def delete_slide(pkg: PowerPointPackage, slide_num: int) -> None:
    """Delete a slide from the presentation."""
    pres = pkg.presentation_xml
    pres_path = pkg.presentation_path

    # Find slide info
    slide_paths = pkg.get_slide_paths()
    target = None
    for num, rid, partname in slide_paths:
        if num == slide_num:
            target = (rid, partname)
            break

    if target is None:
        raise KeyError(f"Slide {slide_num} not found")

    rid, partname = target

    # Remove from sldIdLst
    sld_id_lst = pres.find(qn("p:sldIdLst"), NSMAP)
    if sld_id_lst is not None:
        for sld_id in list(sld_id_lst):
            if sld_id.get(qn("r:id")) == rid:
                sld_id_lst.remove(sld_id)
                break

    # Remove relationship
    pres_rels = pkg.get_rels(pres_path)
    pres_rels.remove(rid)

    # Drop the slide part and its relationships
    pkg.drop_part(partname)

    # Mark dirty and invalidate caches
    pkg.mark_xml_dirty(pres_path)
    pkg.invalidate_caches()


def reorder_slide(pkg: PowerPointPackage, slide_num: int, new_position: int) -> None:
    """Move a slide to a new position."""
    pres = pkg.presentation_xml
    pres_path = pkg.presentation_path

    sld_id_lst = pres.find(qn("p:sldIdLst"), NSMAP)
    if sld_id_lst is None:
        raise ValueError("No slides in presentation")

    sld_ids = list(sld_id_lst)
    if slide_num < 1 or slide_num > len(sld_ids):
        raise KeyError(f"Slide {slide_num} not found")
    if new_position < 1 or new_position > len(sld_ids):
        raise ValueError(f"Invalid position {new_position}")

    # Remove from current position (0-indexed)
    element = sld_ids[slide_num - 1]
    sld_id_lst.remove(element)

    # Insert at new position
    sld_id_lst.insert(new_position - 1, element)

    pkg.mark_xml_dirty(pres_path)
    pkg.invalidate_caches()


def _find_layout_by_name(pkg: PowerPointPackage, layout_name: str | None) -> str:
    """Find layout path by name, or return first layout if name is None."""
    pres_rels = pkg.get_rels(pkg.presentation_path)

    # Get slide master
    master_rid = pres_rels.rId_for_reltype(RT.SLIDE_MASTER)
    if master_rid is None:
        raise ValueError("No slide master found in presentation")

    master_path = pkg.resolve_rel_target(pkg.presentation_path, master_rid)
    master_rels = pkg.get_rels(master_path)

    # Find layouts
    layouts = []
    for rid, rel in master_rels.items():
        if rel.reltype == RT.SLIDE_LAYOUT:
            layout_path = pkg.resolve_rel_target(master_path, rid)
            layouts.append(layout_path)

    if not layouts:
        raise ValueError("No layouts found in slide master")

    if layout_name is None:
        return layouts[0]

    # Search by name
    for layout_path in layouts:
        layout_xml = pkg.get_xml(layout_path)
        cSld = layout_xml.find(qn("p:cSld"), NSMAP)
        if cSld is not None and cSld.get("name") == layout_name:
            return layout_path

    raise KeyError(f"Layout '{layout_name}' not found")


def _create_minimal_slide() -> etree._Element:
    """Create minimal valid slide XML."""
    slide = etree.Element(
        qn("p:sld"),
        nsmap={
            "p": NSMAP["p"],
            "a": NSMAP["a"],
            "r": NSMAP["r"],
        },
    )

    cSld = etree.SubElement(slide, qn("p:cSld"))
    spTree = etree.SubElement(cSld, qn("p:spTree"))

    # Group shape properties (required)
    nvGrpSpPr = etree.SubElement(spTree, qn("p:nvGrpSpPr"))
    etree.SubElement(nvGrpSpPr, qn("p:cNvPr"), id="1", name="")
    etree.SubElement(nvGrpSpPr, qn("p:cNvGrpSpPr"))
    etree.SubElement(nvGrpSpPr, qn("p:nvPr"))

    etree.SubElement(spTree, qn("p:grpSpPr"))

    return slide
