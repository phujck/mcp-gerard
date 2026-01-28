"""Content Control (SDT) operations.

Contains functions for:
- Building list of content controls
- Getting SDT properties (type, value, options, etc.)
- Setting content control values
- Type-specific handlers (checkbox, dropdown, text)

Pure OOXML implementation - works directly with lxml elements.
"""

from __future__ import annotations

from lxml import etree

from mcp_handley_lab.microsoft.word.constants import NSMAP, qn
from mcp_handley_lab.microsoft.word.ops.core import (
    content_hash,
    get_paragraph_text,
    make_block_id,
    mark_dirty,
    paragraph_kind_and_level,
)

# =============================================================================
# Constants
# =============================================================================

# Extended namespace map for SDT content controls (Word 2010/2012 extensions)
_SDT_NSMAP = {
    **NSMAP,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}


# =============================================================================
# SDT Property Helpers
# =============================================================================


def _get_sdt_type(sdt_pr) -> str:
    """Determine the type of content control from its properties."""
    # Check for specific type elements - w14:checkbox for newer Word versions
    if sdt_pr.find("w14:checkbox", namespaces=_SDT_NSMAP) is not None:
        return "checkbox"
    if sdt_pr.find("w:checkbox", namespaces=_SDT_NSMAP) is not None:
        return "checkbox"
    if sdt_pr.find("w:dropDownList", namespaces=_SDT_NSMAP) is not None:
        return "dropdown"
    if sdt_pr.find("w:comboBox", namespaces=_SDT_NSMAP) is not None:
        return "comboBox"
    if sdt_pr.find("w:date", namespaces=_SDT_NSMAP) is not None:
        return "date"
    if sdt_pr.find("w15:color", namespaces=_SDT_NSMAP) is not None:
        return "color"
    if sdt_pr.find("w:richText", namespaces=_SDT_NSMAP) is not None:
        return "richText"
    if sdt_pr.find("w:text", namespaces=_SDT_NSMAP) is not None:
        return "text"
    # Default to text for unrecognized types
    return "text"


def _get_sdt_value(sdt: etree._Element) -> str:
    """Extract the current value from a content control."""
    sdt_content = sdt.find("w:sdtContent", namespaces=NSMAP)
    if sdt_content is None:
        return ""

    # Collect all text from paragraphs and runs
    texts = []
    for p in sdt_content.iter(qn("w:p")):
        para_texts = []
        for t in p.iter(qn("w:t")):
            if t.text:
                para_texts.append(t.text)
        texts.append("".join(para_texts))

    return "\n".join(texts)


def _get_sdt_checked_state(sdt_pr) -> bool | None:
    """Get checkbox checked state from SDT properties."""
    checkbox = sdt_pr.find("w14:checkbox", namespaces=_SDT_NSMAP)
    if checkbox is None:
        # Try w:checkbox for older format
        checkbox = sdt_pr.find("w:checkbox", namespaces=_SDT_NSMAP)

    if checkbox is not None:
        checked = checkbox.find("w14:checked", namespaces=_SDT_NSMAP)
        if checked is None:
            checked = checkbox.find("w:checked", namespaces=_SDT_NSMAP)
        if checked is not None:
            # w14:val attribute
            ns_w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
            val = checked.get(f"{{{ns_w14}}}val") or checked.get(qn("w:val"))
            return val == "1" or val == "true"
        return False

    return None


def _get_sdt_dropdown_options(sdt_pr: etree._Element) -> list[str]:
    """Get dropdown/combobox options from SDT properties."""
    options = []

    # Check dropDownList
    dropdown = sdt_pr.find("w:dropDownList", namespaces=NSMAP)
    if dropdown is None:
        dropdown = sdt_pr.find("w:comboBox", namespaces=NSMAP)

    if dropdown is not None:
        for list_item in dropdown.findall("w:listItem", namespaces=NSMAP):
            display_text = list_item.get(qn("w:displayText"))
            value = list_item.get(qn("w:value"))
            options.append(display_text or value or "")

    return options


def _get_sdt_date_format(sdt_pr: etree._Element) -> str | None:
    """Get date format from SDT properties."""
    date = sdt_pr.find("w:date", namespaces=NSMAP)
    if date is not None:
        date_format = date.find("w:dateFormat", namespaces=NSMAP)
        if date_format is not None:
            return date_format.get(qn("w:val"))
    return None


# =============================================================================
# Block ID Helper
# =============================================================================


def build_block_id_from_element(
    element: etree._Element, block_hash_counts: dict[str, int], para_cache: dict
) -> str:
    """Build block ID from an element using consistent ID system with build_blocks().

    Pure OOXML: Takes w:p element.
    Uses content_hash() for normalization and tracks occurrence by block_type + hash.
    """
    block_type, _ = paragraph_kind_and_level(element)
    text_content = get_paragraph_text(element)
    text_hash = content_hash(text_content)

    # Track occurrence by block_type + hash (same as build_blocks)
    key = f"{block_type}_{text_hash}"
    occurrence = block_hash_counts.get(key, 0)
    block_hash_counts[key] = occurrence + 1

    return make_block_id(block_type, text_content, occurrence)


# =============================================================================
# Main SDT Functions
# =============================================================================


def build_content_controls(pkg) -> list[dict]:
    """Build list of all content controls (SDTs) in the document.

    Args:
        pkg: WordPackage

    Returns list of dicts with: id, tag, alias, type, value, options, checked, date_format, block_id.
    """
    content_controls: list[dict] = []
    block_hash_counts: dict[str, int] = {}

    # Find all SDTs in document body
    body = pkg.body
    if body is None:
        return content_controls

    # Track parent paragraph for block_id
    for sdt in body.iter(qn("w:sdt")):
        sdt_pr = sdt.find("w:sdtPr", namespaces=NSMAP)
        if sdt_pr is None:
            continue

        # Get ID
        id_el = sdt_pr.find("w:id", namespaces=NSMAP)
        sdt_id = int(id_el.get(qn("w:val"))) if id_el is not None else 0

        # Get tag
        tag_el = sdt_pr.find("w:tag", namespaces=NSMAP)
        tag = tag_el.get(qn("w:val")) if tag_el is not None else None

        # Get alias
        alias_el = sdt_pr.find("w:alias", namespaces=NSMAP)
        alias = alias_el.get(qn("w:val")) if alias_el is not None else None

        # Determine type
        sdt_type = _get_sdt_type(sdt_pr)

        # Get value
        value = _get_sdt_value(sdt)

        # Get type-specific info
        checked = _get_sdt_checked_state(sdt_pr) if sdt_type == "checkbox" else None
        options = (
            _get_sdt_dropdown_options(sdt_pr)
            if sdt_type in ("dropdown", "comboBox")
            else []
        )
        date_format = _get_sdt_date_format(sdt_pr) if sdt_type == "date" else None

        # Build block_id (find nearest parent paragraph or use document body)
        parent = sdt.getparent()
        block_id = "document"
        while parent is not None:
            if parent.tag == qn("w:p"):
                # Found parent paragraph
                block_id = build_block_id_from_element(parent, block_hash_counts, {})
                break
            parent = parent.getparent()

        content_controls.append(
            {
                "id": sdt_id,
                "tag": tag,
                "alias": alias,
                "type": sdt_type,
                "value": value,
                "options": options,
                "checked": checked,
                "date_format": date_format,
                "block_id": block_id,
            }
        )

    return content_controls


# =============================================================================
# SDT Value Setters
# =============================================================================


def set_content_control_value(pkg, sdt_id: int, value: str) -> None:
    """Set the value of a content control.

    Args:
        pkg: WordPackage
        sdt_id: Content control ID
        value: New value to set

    For dropdown: value must match one of the options
    For checkbox: value should be "true" or "false"
    For date: value should be ISO date string
    For text: value is the text content
    """
    body = pkg.body
    if body is None:
        raise ValueError("Document has no body")

    # Find SDT with matching ID
    for sdt in body.iter(qn("w:sdt")):
        sdt_pr = sdt.find("w:sdtPr", namespaces=NSMAP)
        if sdt_pr is None:
            continue

        id_el = sdt_pr.find("w:id", namespaces=NSMAP)
        if id_el is None:
            continue

        if int(id_el.get(qn("w:val"))) != sdt_id:
            continue

        # Found matching SDT
        sdt_type = _get_sdt_type(sdt_pr)

        if sdt_type == "checkbox":
            _set_checkbox_value(sdt, sdt_pr, value)
        elif sdt_type in ("dropdown", "comboBox"):
            _set_dropdown_value(sdt, sdt_pr, value)
        else:
            # Text, richText, date, etc.
            _set_text_value(sdt, value)

        # Mark document.xml as modified for WordPackage
        mark_dirty(pkg)
        return

    raise ValueError(f"Content control with ID {sdt_id} not found")


def _set_checkbox_value(
    sdt: etree._Element, sdt_pr: etree._Element, value: str
) -> None:
    """Set checkbox checked state and update displayed content."""
    is_checked = value.lower() in ("true", "1", "yes")
    checked_val = "1" if is_checked else "0"
    ns_w14 = "http://schemas.microsoft.com/office/word/2010/wordml"

    # Find or create checkbox element
    checkbox = sdt_pr.find("w14:checkbox", namespaces=_SDT_NSMAP)
    if checkbox is None:
        checkbox = sdt_pr.find("w:checkbox", namespaces=_SDT_NSMAP)

    if checkbox is not None:
        # Find or create checked element
        checked = checkbox.find("w14:checked", namespaces=_SDT_NSMAP)
        if checked is None:
            checked = checkbox.find("w:checked", namespaces=_SDT_NSMAP)

        if checked is not None:
            # Update existing - use explicit namespace URI
            val_attr_w14 = f"{{{ns_w14}}}val"
            if val_attr_w14 in checked.attrib:
                checked.set(val_attr_w14, checked_val)
            elif qn("w:val") in checked.attrib:
                checked.set(qn("w:val"), checked_val)
            else:
                checked.set(val_attr_w14, checked_val)
        else:
            # Create new checked element
            checked = etree.Element(f"{{{ns_w14}}}checked")
            checked.set(f"{{{ns_w14}}}val", checked_val)
            checkbox.append(checked)

    # Also update the displayed content (checkbox glyph in w:sdtContent)
    # Unicode checkbox characters: checked = ☒ (U+2612), unchecked = ☐ (U+2610)
    display_char = "\u2612" if is_checked else "\u2610"
    _set_text_value(sdt, display_char)


def _set_dropdown_value(
    sdt: etree._Element, sdt_pr: etree._Element, value: str
) -> None:
    """Set dropdown selected value."""
    # Verify value is in options
    options = _get_sdt_dropdown_options(sdt_pr)
    if options and value not in options:
        raise ValueError(f"Value '{value}' not in dropdown options: {options}")

    # Set the text content
    _set_text_value(sdt, value)


def _set_text_value(sdt: etree._Element, value: str) -> None:
    """Set text content of an SDT."""
    sdt_content = sdt.find("w:sdtContent", namespaces=NSMAP)
    if sdt_content is None:
        return

    # Find first paragraph and set its text
    for p in sdt_content.findall("w:p", namespaces=NSMAP):
        # Clear existing runs
        for r in list(p.findall("w:r", namespaces=NSMAP)):
            p.remove(r)

        # Add new run with text
        run = etree.Element(qn("w:r"))
        text = etree.SubElement(run, qn("w:t"))
        text.text = value
        p.append(run)

        return  # Only update first paragraph

    # No paragraph found, create one
    p = etree.Element(qn("w:p"))
    run = etree.SubElement(p, qn("w:r"))
    text = etree.SubElement(run, qn("w:t"))
    text.text = value
    sdt_content.append(p)


# =============================================================================
# SDT Creation
# =============================================================================


def _generate_sdt_id(anchor: etree._Element) -> int:
    """Generate a unique SDT id by scanning all existing w:id values in the document tree."""
    max_id = 0
    root = anchor.getroottree().getroot()
    for sdt in root.iter(qn("w:sdt")):
        sdt_pr = sdt.find("w:sdtPr", namespaces=NSMAP)
        if sdt_pr is not None:
            id_el = sdt_pr.find("w:id", namespaces=NSMAP)
            if id_el is not None:
                val = int(id_el.get(qn("w:val"), "0"))
                max_id = max(max_id, val)
    return max_id + 1


_VALID_SDT_TYPES = {"text", "richText", "dropdown", "comboBox", "checkbox", "date"}


def create_content_control(
    pkg,
    body: etree._Element | None,
    p_el: etree._Element,
    sdt_type: str,
    tag: str | None = None,
    alias: str | None = None,
    placeholder: str = "Click here",
    position: str = "after",
    options: list[str] | None = None,
    checked: bool = False,
    date_format: str = "yyyy-MM-dd",
) -> etree._Element:
    """Create a new block-level content control (SDT).

    Args:
        pkg: WordPackage
        body: Deprecated, ignored. ID generation scans from p_el's root tree.
        p_el: Reference paragraph element for insertion position
        sdt_type: One of "text", "richText", "dropdown", "comboBox", "checkbox", "date"
        tag: Optional tag string
        alias: Optional alias/title string
        placeholder: Placeholder text displayed in the content control
        position: "before" or "after" the reference paragraph
        options: List of option strings (for dropdown/comboBox)
        checked: Initial checked state (for checkbox)
        date_format: Date format string (for date type)

    Returns the created w:sdt element.
    """
    if position not in ("before", "after"):
        raise ValueError(f"position must be 'before' or 'after', got {position!r}")

    if sdt_type not in _VALID_SDT_TYPES:
        raise ValueError(
            f"sdt_type must be one of {sorted(_VALID_SDT_TYPES)}, got {sdt_type!r}"
        )

    # Validate parent is a block-level context
    parent = p_el.getparent()
    if parent is None or parent.tag not in (qn("w:body"), qn("w:tc")):
        raise ValueError(
            "Content controls can only be inserted as siblings of block-level "
            "elements (in w:body or w:tc)"
        )

    sdt_id = _generate_sdt_id(p_el)

    # Build w:sdt
    sdt = etree.Element(qn("w:sdt"))

    # -- w:sdtPr --
    sdt_pr = etree.SubElement(sdt, qn("w:sdtPr"))

    id_el = etree.SubElement(sdt_pr, qn("w:id"))
    id_el.set(qn("w:val"), str(sdt_id))

    if tag:
        tag_el = etree.SubElement(sdt_pr, qn("w:tag"))
        tag_el.set(qn("w:val"), tag)

    if alias:
        alias_el = etree.SubElement(sdt_pr, qn("w:alias"))
        alias_el.set(qn("w:val"), alias)

    # Type-specific sdtPr children
    ns_w14 = "http://schemas.microsoft.com/office/word/2010/wordml"

    if sdt_type == "text":
        etree.SubElement(sdt_pr, qn("w:text"))

    elif sdt_type == "richText":
        etree.SubElement(sdt_pr, qn("w:richText"))

    elif sdt_type == "dropdown":
        dd = etree.SubElement(sdt_pr, qn("w:dropDownList"))
        for opt in options or []:
            item = etree.SubElement(dd, qn("w:listItem"))
            item.set(qn("w:displayText"), opt)
            item.set(qn("w:value"), opt)

    elif sdt_type == "comboBox":
        cb = etree.SubElement(sdt_pr, qn("w:comboBox"))
        for opt in options or []:
            item = etree.SubElement(cb, qn("w:listItem"))
            item.set(qn("w:displayText"), opt)
            item.set(qn("w:value"), opt)

    elif sdt_type == "checkbox":
        checkbox = etree.SubElement(
            sdt_pr, f"{{{ns_w14}}}checkbox", nsmap={"w14": ns_w14}
        )
        checked_el = etree.SubElement(checkbox, f"{{{ns_w14}}}checked")
        checked_el.set(f"{{{ns_w14}}}val", "1" if checked else "0")
        checked_state = etree.SubElement(checkbox, f"{{{ns_w14}}}checkedState")
        checked_state.set(f"{{{ns_w14}}}val", "2612")
        checked_state.set(f"{{{ns_w14}}}font", "MS Gothic")
        unchecked_state = etree.SubElement(checkbox, f"{{{ns_w14}}}uncheckedState")
        unchecked_state.set(f"{{{ns_w14}}}val", "2610")
        unchecked_state.set(f"{{{ns_w14}}}font", "MS Gothic")

    elif sdt_type == "date":
        date_el = etree.SubElement(sdt_pr, qn("w:date"))
        date_fmt = etree.SubElement(date_el, qn("w:dateFormat"))
        date_fmt.set(qn("w:val"), date_format)
        lid = etree.SubElement(date_el, qn("w:lid"))
        lid.set(qn("w:val"), "en-US")
        store = etree.SubElement(date_el, qn("w:storeMappedDataAs"))
        store.set(qn("w:val"), "dateTime")

    # -- w:sdtContent --
    sdt_content = etree.SubElement(sdt, qn("w:sdtContent"))
    p = etree.SubElement(sdt_content, qn("w:p"))
    run = etree.SubElement(p, qn("w:r"))
    t = etree.SubElement(run, qn("w:t"))

    if sdt_type == "checkbox":
        t.text = "\u2612" if checked else "\u2610"
    else:
        t.text = placeholder

    # Insert relative to reference paragraph
    if position == "before":
        p_el.addprevious(sdt)
    else:
        p_el.addnext(sdt)

    mark_dirty(pkg)
    return sdt
