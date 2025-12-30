"""List and numbering operations.

Contains functions for:
- Reading list/numbering properties
- Modifying list levels (promote/demote)
- Restarting numbering
- Removing list formatting
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml.etree import ElementBase as _LxmlElementBase

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph

# Namespace for numbering XPath queries
_NUMBERING_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


# =============================================================================
# Numbering Part Helpers
# =============================================================================


def _get_numbering_part(doc: Document):
    """Get the numbering part from document. Returns None if not present."""
    try:
        return doc.part.numbering_part
    except Exception:
        return None


def _numbering_xpath(element, expr: str) -> list:
    """Execute XPath on element using Word namespace.

    Uses lxml's ElementBase.xpath() directly since python-docx elements
    inherit from lxml but their .xpath() wrapper doesn't support namespaces.
    """
    return _LxmlElementBase.xpath(element, expr, namespaces=_NUMBERING_NS)


def _resolve_abstract_num_id(doc: Document, num_id: int) -> int | None:
    """Resolve num_id to abstractNumId via numbering.xml.

    Returns None if num_id not found.
    """
    numbering_part = _get_numbering_part(doc)
    if numbering_part is None:
        return None

    numbering_el = numbering_part._element
    for num_el in _numbering_xpath(numbering_el, ".//w:num"):
        if num_el.get(qn("w:numId")) == str(num_id):
            abstract_ref = num_el.find(qn("w:abstractNumId"))
            if abstract_ref is not None:
                return int(abstract_ref.get(qn("w:val")))
    return None


def _resolve_level_format(
    doc: Document, abstract_num_id: int, ilvl: int
) -> dict[str, str | int | None]:
    """Get level format info (numFmt, lvlText, start) from abstractNum.

    Returns dict with keys: format_type, level_text, start_value.
    """
    numbering_part = _get_numbering_part(doc)
    if numbering_part is None:
        return {"format_type": None, "level_text": None, "start_value": None}

    numbering_el = numbering_part._element
    for abs_num in _numbering_xpath(numbering_el, ".//w:abstractNum"):
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


def get_list_info(doc: Document, paragraph: Paragraph) -> dict | None:
    """Get list properties for a paragraph.

    Returns None if paragraph is not in a list.
    Returns dict with: num_id, abstract_num_id, level, format_type, start_value, level_text.
    """
    p_el = paragraph._element
    num_pr = p_el.find(qn("w:pPr"))
    if num_pr is None:
        return None

    num_pr_elem = num_pr.find(qn("w:numPr"))
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

    abstract_num_id = _resolve_abstract_num_id(doc, num_id)
    level_info = (
        _resolve_level_format(doc, abstract_num_id, ilvl)
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


def _ensure_pPr(p_el):
    """Ensure paragraph has pPr element, create if needed."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    return pPr


def _ensure_numPr(pPr):
    """Ensure pPr has numPr element, create if needed."""
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = OxmlElement("w:numPr")
        pPr.insert(0, numPr)
    return numPr


def set_list_level(paragraph: Paragraph, level: int) -> None:
    """Set list indentation level (0-8).

    Only works on paragraphs already in a list.
    Raises ValueError if paragraph is not in a list.
    """
    if not 0 <= level <= 8:
        raise ValueError(f"List level must be 0-8, got {level}")

    p_el = paragraph._element
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
        ilvl_el = OxmlElement("w:ilvl")
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(level))


def promote_list_item(paragraph: Paragraph) -> int:
    """Decrease level (move left). Min level is 0. Returns new level."""
    p_el = paragraph._element
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
        ilvl_el = OxmlElement("w:ilvl")
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(new_level))
    return new_level


def demote_list_item(paragraph: Paragraph) -> int:
    """Increase level (move right). Max level is 8. Returns new level."""
    p_el = paragraph._element
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
        ilvl_el = OxmlElement("w:ilvl")
        numPr.insert(0, ilvl_el)
    ilvl_el.set(qn("w:val"), str(new_level))
    return new_level


# =============================================================================
# Numbering Restart
# =============================================================================


def _get_max_num_id(doc: Document) -> int:
    """Get the maximum numId currently in use."""
    numbering_part = _get_numbering_part(doc)
    if numbering_part is None:
        return 0

    max_id = 0
    for num_el in _numbering_xpath(numbering_part._element, ".//w:num"):
        num_id = int(num_el.get(qn("w:numId"), "0"))
        max_id = max(max_id, num_id)
    return max_id


def restart_numbering(doc: Document, paragraph: Paragraph, start_value: int = 1) -> int:
    """Restart numbering from given value.

    Creates a new w:num in numbering.xml referencing the same abstractNum,
    with lvlOverride/startOverride, then updates the paragraph's numId.

    Returns the new numId.
    Raises ValueError if paragraph is not in a list.
    """
    list_info = get_list_info(doc, paragraph)
    if list_info is None:
        raise ValueError("Paragraph is not in a list")

    abstract_num_id = list_info["abstract_num_id"]
    ilvl = list_info["level"]

    if abstract_num_id is None:
        raise ValueError("Cannot determine abstract numbering definition")

    numbering_part = _get_numbering_part(doc)
    if numbering_part is None:
        raise ValueError("Document has no numbering part")

    new_num_id = _get_max_num_id(doc) + 1

    new_num = OxmlElement("w:num")
    new_num.set(qn("w:numId"), str(new_num_id))

    abstract_num_id_el = OxmlElement("w:abstractNumId")
    abstract_num_id_el.set(qn("w:val"), str(abstract_num_id))
    new_num.append(abstract_num_id_el)

    lvl_override = OxmlElement("w:lvlOverride")
    lvl_override.set(qn("w:ilvl"), str(ilvl))

    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), str(start_value))
    lvl_override.append(start_override)

    new_num.append(lvl_override)

    numbering_part._element.append(new_num)

    p_el = paragraph._element
    pPr = p_el.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr"))
    num_id_el = numPr.find(qn("w:numId"))
    num_id_el.set(qn("w:val"), str(new_num_id))

    return new_num_id


# =============================================================================
# Remove List Formatting
# =============================================================================


def remove_list_formatting(paragraph: Paragraph) -> None:
    """Remove list formatting from paragraph (removes w:numPr)."""
    p_el = paragraph._element
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return

    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)
