"""Section and page layout operations.

Contains functions for:
- Page setup reading (margins, orientation, columns, line numbering)
- Setting page margins and orientation
- Adding sections
- Multi-column layout
- Line numbering

Note: These functions still use python-docx Document for section access.
The plan is to migrate to pure OOXML in a future phase.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.section import WD_ORIENT, WD_SECTION
from lxml import etree

from mcp_handley_lab.word.opc.constants import qn

if TYPE_CHECKING:
    from docx import Document

from mcp_handley_lab.word.models import LineNumberingInfo, PageSetupInfo

# EMU per inch for unit conversions
_EMU_PER_INCH = 914400

# =============================================================================
# Constants
# =============================================================================

# Section start type mapping
_SECTION_START_MAP = {
    "new_page": WD_SECTION.NEW_PAGE,
    "continuous": WD_SECTION.CONTINUOUS,
    "even_page": WD_SECTION.EVEN_PAGE,
    "odd_page": WD_SECTION.ODD_PAGE,
    "new_column": WD_SECTION.NEW_COLUMN,
}

# OOXML schema order for w:sectPr children (partial list for insertion)
_SECTPR_ORDER = [
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
]


# =============================================================================
# Page Setup Reading
# =============================================================================


def build_page_setup(doc: Document) -> list[PageSetupInfo]:
    """Build list of PageSetupInfo for all sections."""

    result = []
    for idx, section in enumerate(doc.sections):
        s = section
        sectPr = s._sectPr

        # Extract column settings from w:cols
        columns = 1
        column_spacing = 0.5
        column_separator = False
        cols_el = sectPr.find(qn("w:cols"))
        if cols_el is not None:
            num = cols_el.get(qn("w:num"))
            columns = int(num) if num else 1
            space = cols_el.get(qn("w:space"))
            if space:
                column_spacing = round(int(space) / 1440, 2)  # twips to inches
            sep = cols_el.get(qn("w:sep"))
            column_separator = sep == "1" or sep == "true"

        # Extract line numbering from w:lnNumType
        line_numbering = None
        ln_el = sectPr.find(qn("w:lnNumType"))
        if ln_el is not None:
            restart_map = {
                "newPage": "newPage",
                "newSection": "newSection",
                "continuous": "continuous",
            }
            restart_val = ln_el.get(qn("w:restart")) or "newPage"
            start_val = ln_el.get(qn("w:start")) or "1"
            count_by_val = ln_el.get(qn("w:countBy")) or "1"
            distance_val = ln_el.get(qn("w:distance")) or "720"
            line_numbering = LineNumberingInfo(
                enabled=True,
                restart=restart_map.get(restart_val, "newPage"),
                start=int(start_val),
                count_by=int(count_by_val),
                distance_inches=round(int(distance_val) / 1440, 2),  # twips to inches
            )

        result.append(
            PageSetupInfo(
                section_index=idx,
                orientation="landscape"
                if s.orientation == WD_ORIENT.LANDSCAPE
                else "portrait",
                page_width=round(s.page_width / 914400, 2) if s.page_width else 0.0,
                page_height=round(s.page_height / 914400, 2) if s.page_height else 0.0,
                top_margin=round(s.top_margin / 914400, 2) if s.top_margin else 0.0,
                bottom_margin=round(s.bottom_margin / 914400, 2)
                if s.bottom_margin
                else 0.0,
                left_margin=round(s.left_margin / 914400, 2) if s.left_margin else 0.0,
                right_margin=round(s.right_margin / 914400, 2)
                if s.right_margin
                else 0.0,
                columns=columns,
                column_spacing=column_spacing,
                column_separator=column_separator,
                line_numbering=line_numbering,
            )
        )
    return result


# =============================================================================
# Page Settings
# =============================================================================


def set_page_margins(
    doc: Document,
    section_index: int,
    top: float,
    bottom: float,
    left: float,
    right: float,
) -> None:
    """Set page margins for a section. Values in inches."""
    section = doc.sections[section_index]
    # Set margins in EMU (914400 EMU = 1 inch)
    section.top_margin = int(top * _EMU_PER_INCH)
    section.bottom_margin = int(bottom * _EMU_PER_INCH)
    section.left_margin = int(left * _EMU_PER_INCH)
    section.right_margin = int(right * _EMU_PER_INCH)


def set_page_orientation(doc: Document, section_index: int, orientation: str) -> None:
    """Set page orientation for a section. Valid: 'portrait' or 'landscape'."""
    orient_lower = orientation.lower()
    if orient_lower not in ("portrait", "landscape"):
        raise ValueError(
            f"Invalid orientation '{orientation}'. Valid: ['portrait', 'landscape']"
        )

    section = doc.sections[section_index]
    w, h = section.page_width, section.page_height

    section.orientation = (
        WD_ORIENT.LANDSCAPE if orient_lower == "landscape" else WD_ORIENT.PORTRAIT
    )
    # Swap dimensions if needed - python-docx accepts raw EMU values
    if orient_lower == "landscape" and h > w or orient_lower == "portrait" and w > h:
        section.page_width, section.page_height = h, w


def add_section(doc: Document, start_type: str = "new_page") -> int:
    """Add new section. Returns new section index (0-based).

    start_type: 'new_page', 'continuous', 'even_page', 'odd_page', 'new_column'.
    """
    start_type_lower = start_type.lower()
    if start_type_lower not in _SECTION_START_MAP:
        raise ValueError(
            f"Invalid start_type '{start_type}'. Valid: {list(_SECTION_START_MAP.keys())}"
        )
    doc.add_section(_SECTION_START_MAP[start_type_lower])
    return len(doc.sections) - 1


# =============================================================================
# Section Element Helpers
# =============================================================================


def _insert_sectpr_element(sectPr, element, local_name: str) -> None:
    """Insert element into sectPr at schema-correct position."""
    try:
        target_idx = _SECTPR_ORDER.index(local_name)
    except ValueError:
        # Unknown element, append at end
        sectPr.append(element)
        return

    # Find first child that should come after this element
    for i, child in enumerate(sectPr):
        child_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        try:
            child_idx = _SECTPR_ORDER.index(child_local)
            if child_idx > target_idx:
                sectPr.insert(i, element)
                return
        except ValueError:
            pass

    # No later element found, append
    sectPr.append(element)


# =============================================================================
# Columns and Line Numbering
# =============================================================================


def set_section_columns(
    doc: Document,
    section_index: int,
    num_columns: int,
    spacing_inches: float = 0.5,
    separator: bool = False,
) -> None:
    """Set multi-column layout for a section.

    Args:
        doc: The Document object
        section_index: 0-based section index
        num_columns: Number of columns (1-16)
        spacing_inches: Space between columns in inches
        separator: True to show line between columns
    """
    section = doc.sections[section_index]
    sectPr = section._sectPr

    # Get or create w:cols element
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = etree.Element(qn("w:cols"))
        _insert_sectpr_element(sectPr, cols, "cols")

    # Set attributes
    cols.set(qn("w:num"), str(num_columns))
    cols.set(qn("w:space"), str(int(spacing_inches * 1440)))  # inches to twips
    cols.set(qn("w:sep"), "1" if separator else "0")


def set_line_numbering(
    doc: Document,
    section_index: int,
    enabled: bool = True,
    restart: str = "newPage",
    start: int = 1,
    count_by: int = 1,
    distance_inches: float = 0.5,
) -> None:
    """Enable/configure line numbering for a section.

    Args:
        doc: The Document object
        section_index: 0-based section index
        enabled: True to enable, False to disable line numbering
        restart: When to restart: 'newPage', 'newSection', or 'continuous'
        start: Starting number
        count_by: Show number every N lines
        distance_inches: Distance from margin in inches
    """
    section = doc.sections[section_index]
    sectPr = section._sectPr

    # Find existing w:lnNumType
    lnNumType = sectPr.find(qn("w:lnNumType"))

    if not enabled:
        # Remove if exists
        if lnNumType is not None:
            sectPr.remove(lnNumType)
        return

    # Create if not exists
    if lnNumType is None:
        lnNumType = etree.Element(qn("w:lnNumType"))
        _insert_sectpr_element(sectPr, lnNumType, "lnNumType")

    # Set attributes
    lnNumType.set(qn("w:restart"), restart)
    lnNumType.set(qn("w:start"), str(start))
    lnNumType.set(qn("w:countBy"), str(count_by))
    lnNumType.set(qn("w:distance"), str(int(distance_inches * 1440)))  # inches to twips
