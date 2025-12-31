"""List and numbering operations.

Contains functions for:
- Reading list/numbering properties
- Modifying list levels (promote/demote)
- Restarting numbering
- Removing list formatting

Pure OOXML implementation - works directly with lxml elements.
"""

from __future__ import annotations

from lxml import etree
from lxml.etree import ElementBase as _LxmlElementBase

from mcp_handley_lab.word.opc.constants import qn
from mcp_handley_lab.word.ops.core import mark_dirty

# Namespace for numbering XPath queries
_NUMBERING_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


# =============================================================================
# Duck-Typed Helpers
# =============================================================================


def _get_numbering_xml(pkg) -> etree._Element | None:
    """Get numbering.xml root from WordPackage or Document (duck-typed).

    For Document, returns the live element (modifications persist).
    For WordPackage, returns the parsed element (call mark_xml_dirty after changes).
    """
    if hasattr(pkg, "numbering_xml"):
        return pkg.numbering_xml  # WordPackage
    # Document: access via numbering_part (live element for modification)
    try:
        return pkg.part.numbering_part._element
    except AttributeError:
        return None


def _mark_numbering_dirty(pkg) -> None:
    """Mark numbering.xml as dirty (only needed for WordPackage)."""
    if hasattr(pkg, "mark_xml_dirty"):
        pkg.mark_xml_dirty("/word/numbering.xml")


# =============================================================================
# Numbering Part Helpers
# =============================================================================


def _numbering_xpath(element: etree._Element, expr: str) -> list:
    """Execute XPath on element using Word namespace.

    Duck-typed: Works with lxml elements or python-docx BaseOxmlElement.
    Uses lxml ElementBase.xpath to ensure namespaces param works.
    """
    return _LxmlElementBase.xpath(element, expr, namespaces=_NUMBERING_NS)


def _resolve_abstract_num_id(pkg, num_id: int) -> int | None:
    """Resolve num_id to abstractNumId via numbering.xml.

    Duck-typed: Takes WordPackage or Document.
    Returns None if num_id not found.
    """
    numbering_xml = _get_numbering_xml(pkg)
    if numbering_xml is None:
        return None

    for num_el in _numbering_xpath(numbering_xml, ".//w:num"):
        if num_el.get(qn("w:numId")) == str(num_id):
            abstract_ref = num_el.find(qn("w:abstractNumId"))
            if abstract_ref is not None:
                return int(abstract_ref.get(qn("w:val")))
    return None


def _resolve_level_format(
    pkg, abstract_num_id: int, ilvl: int
) -> dict[str, str | int | None]:
    """Get level format info (numFmt, lvlText, start) from abstractNum.

    Duck-typed: Takes WordPackage or Document.
    Returns dict with keys: format_type, level_text, start_value.
    """
    numbering_xml = _get_numbering_xml(pkg)
    if numbering_xml is None:
        return {"format_type": None, "level_text": None, "start_value": None}

    for abs_num in _numbering_xpath(numbering_xml, ".//w:abstractNum"):
        if abs_num.get(qn("w:abstractNumId")) == str(abstract_num_id):
            for lvl in _numbering_xpath(abs_num, ".//w:lvl"):
                if lvl.get(qn("w:ilvl")) == str(ilvl):
                    num_fmt = lvl.find(qn("w:numFmt"))
                    lvl_text = lvl.find(qn("w:lvlText"))
                    start = lvl.find(qn("w:start"))
                    return {
                        "format_type": num_fmt.get(qn("w:val"))
                        if num_fmt is not None
                        else None,
                        "level_text": lvl_text.get(qn("w:val"))
                        if lvl_text is not None
                        else None,
                        "start_value": int(start.get(qn("w:val")))
                        if start is not None
                        else None,
                    }
    return {"format_type": None, "level_text": None, "start_value": None}


# =============================================================================
# List Info Reading
# =============================================================================


def get_list_info(pkg, p_el: etree._Element) -> dict | None:
    """Get list properties for a paragraph element.

    Duck-typed: Takes WordPackage or Document and w:p element.

    Returns None if paragraph is not in a list.
    Returns dict with: num_id, abstract_num_id, level, format_type, start_value, level_text.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None

    num_pr_elem = pPr.find(qn("w:numPr"))
    if num_pr_elem is None:
        return None

    ilvl_el = num_pr_elem.find(qn("w:ilvl"))
    num_id_el = num_pr_elem.find(qn("w:numId"))

    if num_id_el is None:
        return None

    num_id = int(num_id_el.get(qn("w:val")))
    # numId of 0 means "no list"
    if num_id == 0:
        return None

    ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0

    abstract_num_id = _resolve_abstract_num_id(pkg, num_id)
    level_info = (
        _resolve_level_format(pkg, abstract_num_id, ilvl)
        if abstract_num_id is not None
        else {}
    )

    return {
        "num_id": num_id,
        "abstract_num_id": abstract_num_id,
        "level": ilvl,
        "format_type": level_info.get("format_type"),
        "start_value": level_info.get("start_value"),
        "level_text": level_info.get("level_text"),
    }


# =============================================================================
# List Level Manipulation
# =============================================================================


def _ensure_pPr(p_el: etree._Element) -> etree._Element:
    """Ensure paragraph has pPr element, create if needed.

    Pure OOXML: Works with lxml elements.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p_el.insert(0, pPr)
    return pPr


def _ensure_numPr(pPr: etree._Element) -> etree._Element:
    """Ensure pPr has numPr element, create if needed.

    Pure OOXML: Works with lxml elements.
    """
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = etree.Element(qn("w:numPr"))
        pPr.insert(0, numPr)
    return numPr


def set_list_level(pkg, p_el: etree._Element, level: int) -> None:
    """Set list indentation level (0-8).

    Duck-typed: Takes WordPackage or Document and w:p element.

    Only works on paragraphs already in a list.
    Raises ValueError if paragraph is not in a list.
    """
    if not 0 <= level <= 8:
        raise ValueError(f"List level must be 0-8, got {level}")

    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        raise ValueError("Paragraph is not in a list")

    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        raise ValueError("Paragraph is not in a list")

    num_id_el = numPr.find(qn("w:numId"))
    if num_id_el is None or num_id_el.get(qn("w:val")) == "0":
        raise ValueError("Paragraph is not in a list")

    ilvl_el = numPr.find(qn("w:ilvl"))
    if ilvl_el is None:
        ilvl_el = etree.Element(qn("w:ilvl"))
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(level))
    mark_dirty(pkg)


def promote_list_item(pkg, p_el: etree._Element) -> int:
    """Decrease level (move left). Min level is 0. Returns new level.

    Duck-typed: Takes WordPackage or Document and w:p element.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        raise ValueError("Paragraph is not in a list")

    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        raise ValueError("Paragraph is not in a list")

    num_id_el = numPr.find(qn("w:numId"))
    if num_id_el is None or num_id_el.get(qn("w:val")) == "0":
        raise ValueError("Paragraph is not in a list")

    ilvl_el = numPr.find(qn("w:ilvl"))
    current = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
    new_level = max(0, current - 1)

    if ilvl_el is None:
        ilvl_el = etree.Element(qn("w:ilvl"))
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(new_level))
    mark_dirty(pkg)
    return new_level


def demote_list_item(pkg, p_el: etree._Element) -> int:
    """Increase level (move right). Max level is 8. Returns new level.

    Duck-typed: Takes WordPackage or Document and w:p element.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        raise ValueError("Paragraph is not in a list")

    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        raise ValueError("Paragraph is not in a list")

    num_id_el = numPr.find(qn("w:numId"))
    if num_id_el is None or num_id_el.get(qn("w:val")) == "0":
        raise ValueError("Paragraph is not in a list")

    ilvl_el = numPr.find(qn("w:ilvl"))
    current = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
    new_level = min(8, current + 1)

    if ilvl_el is None:
        ilvl_el = etree.Element(qn("w:ilvl"))
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(new_level))
    mark_dirty(pkg)
    return new_level


# =============================================================================
# Numbering Restart
# =============================================================================


def _get_max_num_id(pkg) -> int:
    """Get the maximum numId currently in use.

    Duck-typed: Takes WordPackage or Document.
    """
    numbering_xml = _get_numbering_xml(pkg)
    if numbering_xml is None:
        return 0

    max_id = 0
    for num_el in _numbering_xpath(numbering_xml, ".//w:num"):
        num_id = int(num_el.get(qn("w:numId"), "0"))
        max_id = max(max_id, num_id)
    return max_id


def restart_numbering(pkg, p_el: etree._Element, start_value: int = 1) -> int:
    """Restart numbering from given value.

    Duck-typed: Takes WordPackage or Document and w:p element.

    Creates a new w:num in numbering.xml referencing the same abstractNum,
    with lvlOverride/startOverride, then updates the paragraph's numId.

    Returns the new numId.
    Raises ValueError if paragraph is not in a list.
    """
    list_info = get_list_info(pkg, p_el)
    if list_info is None:
        raise ValueError("Paragraph is not in a list")

    abstract_num_id = list_info["abstract_num_id"]
    ilvl = list_info["level"]

    if abstract_num_id is None:
        raise ValueError("Cannot determine abstract numbering definition")

    numbering_xml = _get_numbering_xml(pkg)
    if numbering_xml is None:
        raise ValueError("Document has no numbering part")

    new_num_id = _get_max_num_id(pkg) + 1

    new_num = etree.Element(qn("w:num"))
    new_num.set(qn("w:numId"), str(new_num_id))

    abstract_num_id_el = etree.SubElement(new_num, qn("w:abstractNumId"))
    abstract_num_id_el.set(qn("w:val"), str(abstract_num_id))

    lvl_override = etree.SubElement(new_num, qn("w:lvlOverride"))
    lvl_override.set(qn("w:ilvl"), str(ilvl))

    start_override = etree.SubElement(lvl_override, qn("w:startOverride"))
    start_override.set(qn("w:val"), str(start_value))

    numbering_xml.append(new_num)
    _mark_numbering_dirty(pkg)

    pPr = p_el.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr"))
    num_id_el = numPr.find(qn("w:numId"))
    num_id_el.set(qn("w:val"), str(new_num_id))
    mark_dirty(pkg)  # Also mark document.xml dirty

    return new_num_id


# =============================================================================
# Remove List Formatting
# =============================================================================


def remove_list_formatting(pkg, p_el: etree._Element) -> None:
    """Remove list formatting from paragraph (removes w:numPr).

    Duck-typed: Takes WordPackage or Document and w:p element.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return

    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)
        mark_dirty(pkg)
