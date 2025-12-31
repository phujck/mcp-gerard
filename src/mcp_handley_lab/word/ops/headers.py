"""Header, footer, and field operations.

Contains functions for:
- Building/reading headers and footers
- Setting header/footer text
- Appending content to headers/footers
- Clearing headers/footers
- Field insertion (PAGE, NUMPAGES, etc.)
"""

from __future__ import annotations

import json
from typing import TYPE_CHECKING

from docx.shared import Inches
from lxml import etree

from mcp_handley_lab.word.opc.constants import qn

if TYPE_CHECKING:
    from docx import Document

from mcp_handley_lab.word.models import HeaderFooterInfo
from mcp_handley_lab.word.ops.core import (
    _make_run_with,
    make_block_id,
    table_content_for_hash,
)
from mcp_handley_lab.word.ops.tables import populate_table

# EMU per inch for dimension calculations
_EMU_PER_INCH = 914400

# =============================================================================
# Field Insertion
# =============================================================================


def insert_field(
    p_el,
    field_code: str,
    display_text: str = "",
    *,
    uppercase: bool = True,
    placeholder: str = "1",
) -> None:
    """Insert a Word field into a paragraph.

    Pure OOXML: Takes w:p element (or Paragraph wrapper for backwards compat).

    Creates proper OXML field structure with separate runs for each part:
    begin, instruction, separator, result, and end markers.
    Supports any Word field code (PAGE, NUMPAGES, DATE, TIME, AUTHOR, etc.).

    Args:
        p_el: Paragraph element (w:p) or Paragraph wrapper (for backwards compat)
        uppercase: If True (default), uppercases field_code. Set False for case-sensitive
            fields like bookmark names in cross-references.
        placeholder: Default result text shown before field updates.
    """
    code = field_code.strip().upper() if uppercase else field_code

    # Support both lxml element and Paragraph wrapper (backwards compat)
    p = getattr(p_el, "_p", p_el)

    # Run 1: Field begin
    fld_char_begin = etree.Element(qn("w:fldChar"))
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    p.append(_make_run_with(fld_char_begin))

    # Run 2: Field instruction
    instr_text = etree.Element(qn("w:instrText"))
    instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_text.text = f" {code} "
    p.append(_make_run_with(instr_text))

    # Run 3: Field separator
    fld_char_sep = etree.Element(qn("w:fldChar"))
    fld_char_sep.set(qn("w:fldCharType"), "separate")
    p.append(_make_run_with(fld_char_sep))

    # Run 4: Result text (placeholder shown before field updates)
    text_elem = etree.Element(qn("w:t"))
    text_elem.text = display_text or placeholder
    p.append(_make_run_with(text_elem))

    # Run 5: Field end
    fld_char_end = etree.Element(qn("w:fldChar"))
    fld_char_end.set(qn("w:fldCharType"), "end")
    p.append(_make_run_with(fld_char_end))


def insert_page_x_of_y(
    doc: Document, section_index: int, location: str = "footer"
) -> None:
    """Insert 'Page X of Y' into header or footer.

    Args:
        doc: Document to modify
        section_index: 0-based section index
        location: 'header' or 'footer'
    """
    section = doc.sections[section_index]  # Let IndexError propagate
    hf = {"footer": section.footer, "header": section.header}[
        location
    ]  # KeyError if invalid
    hf.is_linked_to_previous = False
    p = hf.add_paragraph()
    p.add_run("Page ")
    insert_field(p, "PAGE")
    p.add_run(" of ")
    insert_field(p, "NUMPAGES")


# =============================================================================
# Header/Footer Reading
# =============================================================================


def _hf_text(hf) -> str | None:
    """Extract header/footer text. Returns None if linked to previous section."""
    return (
        None if hf.is_linked_to_previous else "\n".join(p.text for p in hf.paragraphs)
    )


def build_headers_footers(doc: Document) -> list[HeaderFooterInfo]:
    """Build list of HeaderFooterInfo for all sections."""
    result = []
    for idx, section in enumerate(doc.sections):
        hdr, ftr = section.header, section.footer
        info = HeaderFooterInfo(
            section_index=idx,
            header_text=_hf_text(hdr),
            footer_text=_hf_text(ftr),
            header_is_linked=hdr.is_linked_to_previous,
            footer_is_linked=ftr.is_linked_to_previous,
            has_different_first_page=section.different_first_page_header_footer,
            has_different_odd_even=doc.settings.odd_and_even_pages_header_footer,
        )
        if section.different_first_page_header_footer:
            fp_hdr, fp_ftr = section.first_page_header, section.first_page_footer
            info.first_page_header_text = _hf_text(fp_hdr)
            info.first_page_footer_text = _hf_text(fp_ftr)
        if doc.settings.odd_and_even_pages_header_footer:
            ev_hdr, ev_ftr = section.even_page_header, section.even_page_footer
            info.even_page_header_text = _hf_text(ev_hdr)
            info.even_page_footer_text = _hf_text(ev_ftr)
        result.append(info)
    return result


# =============================================================================
# Header/Footer Modification
# =============================================================================


def set_header_footer_text(
    doc: Document, section_index: int, text: str, location: str
) -> None:
    """Set header/footer text. Handles all types via location attribute name."""
    section = doc.sections[section_index]
    if location.startswith("first_page_"):
        section.different_first_page_header_footer = True
    elif location.startswith("even_page_"):
        doc.settings.odd_and_even_pages_header_footer = True
    hf = getattr(section, location)
    for p in list(hf.paragraphs):
        p._element.getparent().remove(p._element)
    hf.add_paragraph(text)


def append_to_header_footer(
    doc: Document,
    section_index: int,
    content_type: str,
    content_data: str,
    location: str,
) -> str:
    """Append paragraph or table to header/footer. Returns element_id."""
    hf = getattr(doc.sections[section_index], location)
    if content_type == "paragraph":
        p = hf.add_paragraph(content_data)
        occurrence = sum(1 for para in hf.paragraphs if para.text == p.text) - 1
        return make_block_id("paragraph", p.text, occurrence)
    if content_type == "table":
        table_data = json.loads(content_data)
        rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
        tbl = hf.add_table(rows=rows, cols=cols, width=Inches(6))
        populate_table(tbl._tbl, table_data)  # Pass element, not wrapper
        return make_block_id("table", table_content_for_hash(tbl._tbl), 0)
    raise ValueError(f"Unknown content_type: {content_type}")


def clear_header_footer(doc: Document, section_index: int, location: str) -> None:
    """Clear header/footer content. Unlinks from previous section first."""
    hf = getattr(doc.sections[section_index], location)
    hf.is_linked_to_previous = False
    hf_el = hf._element
    for child in list(hf_el):
        hf_el.remove(child)
    hf.add_paragraph("")  # Word requires at least one paragraph
