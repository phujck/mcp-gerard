"""Footnote and endnote operations.

Contains functions for:
- Reading footnotes/endnotes from documents
- Adding footnotes/endnotes to specific blocks
- Deleting footnotes/endnotes
- Managing footnote/endnote XML parts via zipfile operations
"""

from __future__ import annotations

import os
import zipfile
from typing import TYPE_CHECKING

from docx import Document
from lxml import etree
from lxml.etree import ElementBase as _LxmlElementBase

from mcp_handley_lab.word.opc.constants import qn

if TYPE_CHECKING:
    pass

from mcp_handley_lab.word.ops.core import (
    count_occurrence,
    find_paragraph_by_id,
    get_paragraph_text,
    make_block_id,
    paragraph_kind_and_level,
)

# =============================================================================
# Constants
# =============================================================================

# Namespaces for footnote operations
_FN_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_FN_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_FN_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_FN_XML_NS = "http://www.w3.org/XML/1998/namespace"

# Reserved footnote/endnote IDs (separators only, user notes use 1+)
_RESERVED_NOTE_IDS = {-1, 0}


# =============================================================================
# XPath Helper (duck-typed for lxml compatibility)
# =============================================================================


def _fn_xpath(element: etree._Element, expr: str, ns: dict) -> list:
    """XPath using ElementBase.xpath for namespace compatibility.

    Uses lxml ElementBase.xpath to ensure namespaces parameter works
    correctly with both raw lxml elements and python-docx BaseOxmlElement.
    """
    return _LxmlElementBase.xpath(element, expr, namespaces=ns)


# =============================================================================
# Reading Footnotes/Endnotes
# =============================================================================


def build_footnotes(pkg) -> list[dict]:
    """List all footnotes and endnotes with their content.

    Reads from word/footnotes.xml and word/endnotes.xml, matching references
    in the document body. Returns list of dicts with id, type, text, block_id.

    Args:
        pkg: WordPackage or python-docx Document (duck-typed)
    """
    result = []
    ns = {"w": _FN_W_NS}

    # Get document element (duck-typed for WordPackage vs Document)
    if hasattr(pkg, "document_xml"):
        doc_element = pkg.document_xml
    else:
        doc_element = pkg.element

    # Build map of reference locations in document (id -> block_id)
    ref_locations: dict[tuple[str, int], str] = {}  # (type, id) -> block_id

    # Find footnote references in document body
    for fn_ref in _fn_xpath(doc_element, ".//w:footnoteReference", ns):
        fn_id = fn_ref.get(qn("w:id"))
        if fn_id:
            # Find containing block
            parent = fn_ref.getparent()
            while parent is not None:
                if parent.tag == qn("w:p"):
                    kind, level = paragraph_kind_and_level(parent)
                    text = get_paragraph_text(parent)
                    occurrence = count_occurrence(pkg, kind, text, parent)
                    block_id = make_block_id(kind, text, occurrence)
                    ref_locations[("footnote", int(fn_id))] = block_id
                    break
                parent = parent.getparent()

    # Find endnote references
    for en_ref in _fn_xpath(doc_element, ".//w:endnoteReference", ns):
        en_id = en_ref.get(qn("w:id"))
        if en_id:
            parent = en_ref.getparent()
            while parent is not None:
                if parent.tag == qn("w:p"):
                    kind, level = paragraph_kind_and_level(parent)
                    text = get_paragraph_text(parent)
                    occurrence = count_occurrence(pkg, kind, text, parent)
                    block_id = make_block_id(kind, text, occurrence)
                    ref_locations[("endnote", int(en_id))] = block_id
                    break
                parent = parent.getparent()

    # Read footnotes/endnotes (duck-typed for WordPackage vs Document)
    if hasattr(pkg, "footnotes_xml"):
        # WordPackage: use direct property access
        fn_root = pkg.footnotes_xml
        if fn_root is not None:
            for fn in _fn_xpath(fn_root, "//w:footnote", ns):
                fn_id_str = fn.get(f"{{{_FN_W_NS}}}id")
                if fn_id_str:
                    fn_id = int(fn_id_str)
                    if fn_id in _RESERVED_NOTE_IDS:
                        continue  # Skip separators
                    text_parts = _fn_xpath(fn, ".//w:t/text()", ns)
                    text = "".join(text_parts).strip()
                    block_id = ref_locations.get(("footnote", fn_id), "")
                    result.append(
                        {
                            "id": fn_id,
                            "type": "footnote",
                            "text": text,
                            "block_id": block_id,
                        }
                    )
        en_root = pkg.endnotes_xml
        if en_root is not None:
            for en in _fn_xpath(en_root, "//w:endnote", ns):
                en_id_str = en.get(f"{{{_FN_W_NS}}}id")
                if en_id_str:
                    en_id = int(en_id_str)
                    if en_id in _RESERVED_NOTE_IDS:
                        continue  # Skip separators
                    text_parts = _fn_xpath(en, ".//w:t/text()", ns)
                    text = "".join(text_parts).strip()
                    block_id = ref_locations.get(("endnote", en_id), "")
                    result.append(
                        {
                            "id": en_id,
                            "type": "endnote",
                            "text": text,
                            "block_id": block_id,
                        }
                    )
    else:
        # python-docx Document: iterate package parts
        for part in pkg.part.package.iter_parts():
            partname = part.partname
            if partname == "/word/footnotes.xml":
                fn_root = etree.fromstring(part.blob)
                for fn in _fn_xpath(fn_root, "//w:footnote", ns):
                    fn_id_str = fn.get(f"{{{_FN_W_NS}}}id")
                    if fn_id_str:
                        fn_id = int(fn_id_str)
                        if fn_id in _RESERVED_NOTE_IDS:
                            continue  # Skip separators
                        text_parts = _fn_xpath(fn, ".//w:t/text()", ns)
                        text = "".join(text_parts).strip()
                        block_id = ref_locations.get(("footnote", fn_id), "")
                        result.append(
                            {
                                "id": fn_id,
                                "type": "footnote",
                                "text": text,
                                "block_id": block_id,
                            }
                        )
            elif partname == "/word/endnotes.xml":
                en_root = etree.fromstring(part.blob)
                for en in _fn_xpath(en_root, "//w:endnote", ns):
                    en_id_str = en.get(f"{{{_FN_W_NS}}}id")
                    if en_id_str:
                        en_id = int(en_id_str)
                        if en_id in _RESERVED_NOTE_IDS:
                            continue  # Skip separators
                        text_parts = _fn_xpath(en, ".//w:t/text()", ns)
                        text = "".join(text_parts).strip()
                        block_id = ref_locations.get(("endnote", en_id), "")
                        result.append(
                            {
                                "id": en_id,
                                "type": "endnote",
                                "text": text,
                                "block_id": block_id,
                            }
                        )

    return result


# =============================================================================
# Helper Functions for Zipfile Operations
# =============================================================================


def _get_safe_note_id(notes_root, ns: dict) -> int:
    """Get a safe footnote/endnote ID avoiding reserved values."""
    used_ids: set[int] = set()
    for fn in _fn_xpath(notes_root, "//w:footnote | //w:endnote", ns):
        fn_id = fn.get(f"{{{_FN_W_NS}}}id")
        if fn_id:
            try:
                used_ids.add(int(fn_id))
            except ValueError:
                pass

    candidate_id = 1  # User notes start at 1, reserved are -1 and 0
    while candidate_id in used_ids or candidate_id in _RESERVED_NOTE_IDS:
        candidate_id += 1
    return candidate_id


def _ensure_note_content_types(ct_xml: bytes, note_type: str) -> bytes:
    """Ensure content types include footnotes/endnotes."""
    ct_root = etree.fromstring(ct_xml)
    ns = {"ct": _FN_CT_NS}

    part_name = f"/word/{note_type}s.xml"
    if note_type == "footnote":
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    else:
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"

    existing = _fn_xpath(ct_root, f"//ct:Override[@PartName='{part_name}']", ns)
    if existing:
        return ct_xml

    override = etree.Element(
        f"{{{_FN_CT_NS}}}Override",
        PartName=part_name,
        ContentType=content_type,
    )
    ct_root.append(override)
    return etree.tostring(
        ct_root, encoding="UTF-8", xml_declaration=True, standalone="yes"
    )


def _ensure_note_relationship(rels_xml: bytes, note_type: str) -> bytes:
    """Ensure document relationships include footnotes/endnotes."""
    rels_root = etree.fromstring(rels_xml)
    ns = {"r": _FN_REL_NS}

    rel_type = f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{note_type}s"
    existing = _fn_xpath(
        rels_root, f"//r:Relationship[contains(@Type, '{note_type}s')]", ns
    )
    if existing:
        return rels_xml

    # Generate unique rId
    all_rels = _fn_xpath(rels_root, "//r:Relationship", ns)
    existing_ids = {rel.get("Id") for rel in all_rels if rel.get("Id")}
    rid_num = 1
    while f"rId{rid_num}" in existing_ids:
        rid_num += 1

    rel = etree.Element(
        f"{{{_FN_REL_NS}}}Relationship",
        Id=f"rId{rid_num}",
        Type=rel_type,
        Target=f"{note_type}s.xml",
    )
    rels_root.append(rel)
    return etree.tostring(
        rels_root, encoding="UTF-8", xml_declaration=True, standalone="yes"
    )


def _create_minimal_notes_xml(note_type: str) -> bytes:
    """Create minimal footnotes.xml or endnotes.xml with separators."""
    tag = "footnotes" if note_type == "footnote" else "endnotes"
    item = "footnote" if note_type == "footnote" else "endnote"
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:{tag} xmlns:w="{_FN_W_NS}">
    <w:{item} w:type="separator" w:id="-1">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:separator/>
            </w:r>
        </w:p>
    </w:{item}>
    <w:{item} w:type="continuationSeparator" w:id="0">
        <w:p>
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
            <w:r>
                <w:continuationSeparator/>
            </w:r>
        </w:p>
    </w:{item}>
</w:{tag}>'''
    return xml.encode("utf-8")


def _ensure_note_styles(styles_root, note_type: str) -> None:
    """Ensure footnote/endnote styles exist."""
    ns = {"w": _FN_W_NS}
    ref_style_id = f"{note_type.title()}Reference"
    text_style_id = f"{note_type.title()}Text"

    # Check for reference style
    ref_style = _fn_xpath(styles_root, f'//w:style[@w:styleId="{ref_style_id}"]', ns)
    if not ref_style:
        style = etree.Element(
            f"{{{_FN_W_NS}}}style",
            attrib={
                f"{{{_FN_W_NS}}}type": "character",
                f"{{{_FN_W_NS}}}styleId": ref_style_id,
            },
        )
        name = etree.SubElement(style, f"{{{_FN_W_NS}}}name")
        name.set(f"{{{_FN_W_NS}}}val", f"{note_type} reference")
        base = etree.SubElement(style, f"{{{_FN_W_NS}}}basedOn")
        base.set(f"{{{_FN_W_NS}}}val", "DefaultParagraphFont")
        rPr = etree.SubElement(style, f"{{{_FN_W_NS}}}rPr")
        vert_align = etree.SubElement(rPr, f"{{{_FN_W_NS}}}vertAlign")
        vert_align.set(f"{{{_FN_W_NS}}}val", "superscript")
        styles_root.append(style)

    # Check for text style
    text_style = _fn_xpath(styles_root, f'//w:style[@w:styleId="{text_style_id}"]', ns)
    if not text_style:
        style = etree.Element(
            f"{{{_FN_W_NS}}}style",
            attrib={
                f"{{{_FN_W_NS}}}type": "paragraph",
                f"{{{_FN_W_NS}}}styleId": text_style_id,
            },
        )
        name = etree.SubElement(style, f"{{{_FN_W_NS}}}name")
        name.set(f"{{{_FN_W_NS}}}val", f"{note_type} text")
        base = etree.SubElement(style, f"{{{_FN_W_NS}}}basedOn")
        base.set(f"{{{_FN_W_NS}}}val", "Normal")
        pPr = etree.SubElement(style, f"{{{_FN_W_NS}}}pPr")
        sz = etree.SubElement(pPr, f"{{{_FN_W_NS}}}sz")
        sz.set(f"{{{_FN_W_NS}}}val", "20")  # 10pt
        styles_root.append(style)


# =============================================================================
# Add/Delete Footnotes (Zipfile-based)
# =============================================================================


def add_footnote(
    doc_path: str,
    target_id: str,
    text: str,
    note_type: str = "footnote",
    position: str = "after",
) -> int:
    """Add a footnote or endnote to a document.

    Args:
        doc_path: Path to the Word document
        target_id: Block ID where the reference should be placed
        text: Content of the footnote/endnote
        note_type: "footnote" or "endnote"
        position: "after" (end of paragraph) or "before" (start)

    Returns:
        The ID of the new footnote/endnote

    Raises:
        ValueError: If target not found or invalid location
        FileNotFoundError: If document not found
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")

    ns = {"w": _FN_W_NS}
    notes_file = f"word/{note_type}s.xml"
    note_tag = note_type

    # Read document parts
    doc_parts: dict[str, bytes] = {}
    with zipfile.ZipFile(doc_path, "r") as zf:
        doc_parts["document"] = zf.read("word/document.xml")
        doc_parts["content_types"] = zf.read("[Content_Types].xml")
        doc_parts["document_rels"] = zf.read("word/_rels/document.xml.rels")

        if notes_file in zf.namelist():
            doc_parts["notes"] = zf.read(notes_file)
        else:
            doc_parts["notes"] = _create_minimal_notes_xml(note_type)

        if "word/styles.xml" in zf.namelist():
            doc_parts["styles"] = zf.read("word/styles.xml")
        else:
            doc_parts["styles"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{_FN_W_NS}"/>'.encode()
            )

    # Parse XML
    doc_root = etree.fromstring(doc_parts["document"])
    notes_root = etree.fromstring(doc_parts["notes"])
    styles_root = etree.fromstring(doc_parts["styles"])

    # Load document with python-docx to find target paragraph
    temp_doc = Document(doc_path)

    # Use find_paragraph_by_id which returns lxml element (or None)
    target_para_el = find_paragraph_by_id(temp_doc, target_id)
    if target_para_el is None:
        raise ValueError(f"Target block not found: {target_id}")

    # Get all paragraphs from both python-docx body and raw XML
    # Use index-based matching (more robust than text matching for duplicates)
    all_paras_docx = list(temp_doc.element.body.iter(qn("w:p")))
    all_paras_xml = _fn_xpath(doc_root, "//w:body//w:p", ns)

    # Find index of target in python-docx element tree
    try:
        target_idx = all_paras_docx.index(target_para_el)  # Element, not wrapper
    except ValueError:
        raise ValueError(f"Target element not found in document structure: {target_id}")

    if target_idx >= len(all_paras_xml):
        raise ValueError(f"Paragraph index mismatch: {target_idx}")

    target_para = all_paras_xml[target_idx]

    # Validate location (not in header/footer)
    parent = target_para.getparent()
    while parent is not None:
        if parent.tag in [f"{{{_FN_W_NS}}}hdr", f"{{{_FN_W_NS}}}ftr"]:
            raise ValueError("Cannot add footnote/endnote in header/footer")
        parent = parent.getparent()

    # Get safe note ID
    note_id = _get_safe_note_id(notes_root, ns)

    # Find insertion position in paragraph
    runs = _fn_xpath(target_para, ".//w:r", ns)
    if position == "after" and runs:
        last_run = runs[-1]
        insert_pos = list(target_para).index(last_run) + 1
    elif position == "before" and runs:
        first_run = runs[0]
        insert_pos = list(target_para).index(first_run)
    else:
        insert_pos = len(target_para)

    # Create footnote reference run in document
    ref_run = etree.Element(f"{{{_FN_W_NS}}}r")
    rPr = etree.SubElement(ref_run, f"{{{_FN_W_NS}}}rPr")
    rStyle = etree.SubElement(rPr, f"{{{_FN_W_NS}}}rStyle")
    rStyle.set(f"{{{_FN_W_NS}}}val", f"{note_type.title()}Reference")
    fn_ref = etree.SubElement(ref_run, f"{{{_FN_W_NS}}}{note_type}Reference")
    fn_ref.set(f"{{{_FN_W_NS}}}id", str(note_id))
    target_para.insert(insert_pos, ref_run)

    # Create footnote content in notes file
    new_note = etree.Element(
        f"{{{_FN_W_NS}}}{note_tag}",
        attrib={f"{{{_FN_W_NS}}}id": str(note_id)},
    )

    # Add paragraph with content
    fn_para = etree.SubElement(new_note, f"{{{_FN_W_NS}}}p")
    pPr = etree.SubElement(fn_para, f"{{{_FN_W_NS}}}pPr")
    pStyle = etree.SubElement(pPr, f"{{{_FN_W_NS}}}pStyle")
    pStyle.set(f"{{{_FN_W_NS}}}val", f"{note_type.title()}Text")

    # Add footnote reference marker
    marker_run = etree.SubElement(fn_para, f"{{{_FN_W_NS}}}r")
    marker_rPr = etree.SubElement(marker_run, f"{{{_FN_W_NS}}}rPr")
    marker_rStyle = etree.SubElement(marker_rPr, f"{{{_FN_W_NS}}}rStyle")
    marker_rStyle.set(f"{{{_FN_W_NS}}}val", f"{note_type.title()}Reference")
    etree.SubElement(marker_run, f"{{{_FN_W_NS}}}{note_type}Ref")

    # Add space
    space_run = etree.SubElement(fn_para, f"{{{_FN_W_NS}}}r")
    space_text = etree.SubElement(space_run, f"{{{_FN_W_NS}}}t")
    space_text.set(f"{{{_FN_XML_NS}}}space", "preserve")
    space_text.text = " "

    # Add content text
    text_run = etree.SubElement(fn_para, f"{{{_FN_W_NS}}}r")
    text_elem = etree.SubElement(text_run, f"{{{_FN_W_NS}}}t")
    text_elem.text = text

    notes_root.append(new_note)

    # Ensure styles exist
    _ensure_note_styles(styles_root, note_type)

    # Ensure content types and relationships
    ct_xml = _ensure_note_content_types(doc_parts["content_types"], note_type)
    rels_xml = _ensure_note_relationship(doc_parts["document_rels"], note_type)

    # Write modified document
    temp_path = doc_path + ".tmp"
    with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        with zipfile.ZipFile(doc_path, "r") as zin:
            for item in zin.infolist():
                if item.filename not in [
                    "word/document.xml",
                    notes_file,
                    "word/styles.xml",
                    "[Content_Types].xml",
                    "word/_rels/document.xml.rels",
                ]:
                    zout.writestr(item, zin.read(item.filename))

        zout.writestr(
            "word/document.xml",
            etree.tostring(
                doc_root, encoding="UTF-8", xml_declaration=True, standalone="yes"
            ),
        )
        zout.writestr(
            notes_file,
            etree.tostring(
                notes_root, encoding="UTF-8", xml_declaration=True, standalone="yes"
            ),
        )
        zout.writestr(
            "word/styles.xml",
            etree.tostring(
                styles_root, encoding="UTF-8", xml_declaration=True, standalone="yes"
            ),
        )
        zout.writestr("[Content_Types].xml", ct_xml)
        zout.writestr("word/_rels/document.xml.rels", rels_xml)

    os.replace(temp_path, doc_path)
    return note_id


def delete_footnote(doc_path: str, note_id: int, note_type: str = "footnote") -> None:
    """Delete a footnote or endnote from a document.

    Args:
        doc_path: Path to the Word document
        note_id: ID of the footnote/endnote to delete
        note_type: "footnote" or "endnote"

    Raises:
        ValueError: If note not found
        FileNotFoundError: If document not found
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")

    ns = {"w": _FN_W_NS}
    notes_file = f"word/{note_type}s.xml"

    # Read document parts
    with zipfile.ZipFile(doc_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
        if notes_file not in zf.namelist():
            raise ValueError(f"No {note_type}s in document")
        notes_xml = zf.read(notes_file)

    # Parse XML
    doc_root = etree.fromstring(doc_xml)
    notes_root = etree.fromstring(notes_xml)

    # Remove reference from document
    refs_removed = 0
    for fn_ref in _fn_xpath(
        doc_root, f"//w:{note_type}Reference[@w:id='{note_id}']", ns
    ):
        run = fn_ref.getparent()
        if run is not None and run.tag == f"{{{_FN_W_NS}}}r":
            para = run.getparent()
            if para is not None:
                para.remove(run)
                refs_removed += 1

    if refs_removed == 0:
        raise ValueError(f"{note_type.title()} {note_id} not found")

    # Remove note content
    for fn in _fn_xpath(notes_root, f"//w:{note_type}[@w:id='{note_id}']", ns):
        notes_root.remove(fn)

    # Write modified document
    temp_path = doc_path + ".tmp"
    with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        with zipfile.ZipFile(doc_path, "r") as zin:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(
                        item,
                        etree.tostring(
                            doc_root,
                            encoding="UTF-8",
                            xml_declaration=True,
                            standalone="yes",
                        ),
                    )
                elif item.filename == notes_file:
                    zout.writestr(
                        item,
                        etree.tostring(
                            notes_root,
                            encoding="UTF-8",
                            xml_declaration=True,
                            standalone="yes",
                        ),
                    )
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(temp_path, doc_path)
