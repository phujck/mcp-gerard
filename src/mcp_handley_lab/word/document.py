"""Python-docx wrapper functions for Word document manipulation."""

import hashlib
import json
import re
from collections.abc import Iterator

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from mcp_handley_lab.word.models import (
    Block,
    CellInfo,
    CommentInfo,
    DocumentMeta,
    HeaderFooterInfo,
    PageSetupInfo,
    RunInfo,
)

_HEADING_RE = re.compile(r"^Heading ([1-9])$")
_ID_RE = re.compile(r"^(paragraph|heading[1-9]|table)_([0-9a-f]{8})_(\d+)$")


def content_hash(text: str) -> str:
    """8-char SHA256 of normalized text for content-addressable IDs."""
    normalized = re.sub(r"\s+", " ", text.strip().lower())
    return hashlib.sha256(normalized.encode()).hexdigest()[:8]


def iter_body_blocks(
    doc: Document,
) -> Iterator[tuple[str, Paragraph | Table, CT_P | CT_Tbl]]:
    """Yield (block_kind, block_obj, element) in true document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield ("paragraph", Paragraph(child, doc), child)
        elif isinstance(child, CT_Tbl):
            yield ("table", Table(child, doc), child)


def paragraph_kind_and_level(p: Paragraph) -> tuple[str, int]:
    """Detect if paragraph is a heading and return (kind, level).

    For content-hash IDs, kind includes level: 'heading1', 'heading2', etc.
    """
    style_name = p.style.name if p.style else "Normal"
    m = _HEADING_RE.match(style_name)
    if m:
        level = int(m.group(1))
        return (f"heading{level}", level)
    return ("paragraph", 0)


def table_content_for_hash(table: Table) -> str:
    """Get full table content as canonical string for hashing."""
    return json.dumps(
        [[c.text for c in r.cells] for r in table.rows],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def table_to_markdown(
    table: Table, max_chars: int = 500, max_rows: int = 20, max_cols: int = 10
) -> tuple[str, int, int]:
    """Convert table to markdown preview with truncation."""
    rows, cols = len(table.rows), len(table.columns)
    r_lim, c_lim = min(rows, max_rows), min(cols, max_cols)

    grid = [
        [
            table.cell(r, c).text.strip().replace("|", "\\|").replace("\n", "<br>")
            for c in range(c_lim)
        ]
        for r in range(r_lim)
    ]

    header = grid[0]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in grid[1:])

    if rows > r_lim:
        lines.append(f"... ({rows - r_lim} more rows)")
    if cols > c_lim:
        lines.append(f"... ({cols - c_lim} more cols)")

    md = "\n".join(lines)
    if len(md) > max_chars:
        md = md[:max_chars] + "\n... (truncated)"
    return md, rows, cols


def get_document_meta(doc: Document) -> DocumentMeta:
    """Extract document metadata from core properties."""
    cp = doc.core_properties
    return DocumentMeta(
        title=cp.title or "",
        author=cp.author or "",
        created=cp.created.isoformat() if cp.created else "",
        modified=cp.modified.isoformat() if cp.modified else "",
        revision=cp.revision or 0,
        sections=len(doc.sections),
    )


def build_blocks(
    doc: Document,
    offset: int = 0,
    limit: int = 50,
    heading_only: bool = False,
    search_query: str = "",
) -> tuple[list[Block], int]:
    """Build list of Block objects from document body.

    Uses content-hash IDs: {type}_{hash}_{occurrence}
    """
    blocks = []
    matched = 0
    query_lower = search_query.lower() if search_query else ""
    hash_counts: dict[str, int] = {}  # {type_hash: count} for occurrence tracking

    for kind, obj, _el in iter_body_blocks(doc):
        if kind == "paragraph":
            block_type, level = paragraph_kind_and_level(obj)
            text = obj.text or ""
            style = obj.style.name if obj.style else "Normal"

            # Track occurrence for this type+hash combination
            hash_key = f"{block_type}_{content_hash(text)}"
            occurrence = hash_counts.get(hash_key, 0)
            hash_counts[hash_key] = occurrence + 1

            if heading_only and not block_type.startswith("heading"):
                continue
            if query_lower and query_lower not in text.lower():
                continue

            if matched >= offset and len(blocks) < limit:
                blocks.append(
                    Block(
                        id=f"{block_type}_{content_hash(text)}_{occurrence}",
                        type=block_type,
                        text=text,
                        style=style,
                        level=level,
                    )
                )
            matched += 1

        elif kind == "table":
            md, rows, cols = table_to_markdown(obj)
            style = obj.style.name if obj.style else ""
            # Use full content for hash (not truncated markdown)
            table_content = table_content_for_hash(obj)

            # Track occurrence for this type+hash combination
            hash_key = f"table_{content_hash(table_content)}"
            occurrence = hash_counts.get(hash_key, 0)
            hash_counts[hash_key] = occurrence + 1

            if heading_only:
                continue
            # Search full table content, not truncated markdown preview
            if query_lower and query_lower not in table_content.lower():
                continue

            if matched >= offset and len(blocks) < limit:
                blocks.append(
                    Block(
                        id=f"table_{content_hash(table_content)}_{occurrence}",
                        type="table",
                        text=md,  # Keep markdown for display
                        style=style,
                        rows=rows,
                        cols=cols,
                    )
                )
            matched += 1

    return blocks, matched


def resolve_target(
    doc: Document, target_id: str
) -> tuple[str, Paragraph | Table, CT_P | CT_Tbl, int]:
    """Resolve target_id to (block_type, obj, element, occurrence).

    Uses content-hash IDs: {type}_{hash}_{occurrence}
    Searches all blocks for matching type+hash, then skips to Nth occurrence.
    """
    m = _ID_RE.match(target_id)
    if not m:
        raise ValueError(f"Bad block id: {target_id}")
    target_type, target_hash, occurrence_str = m.groups()
    target_occurrence = int(occurrence_str)

    occurrence_count = 0
    for kind, obj, el in iter_body_blocks(doc):
        if kind == "table":
            if target_type == "table":
                table_content = table_content_for_hash(obj)
                if content_hash(table_content) == target_hash:
                    if occurrence_count == target_occurrence:
                        return "table", obj, el, occurrence_count
                    occurrence_count += 1
        else:
            block_type, _ = paragraph_kind_and_level(obj)
            if target_type == block_type:
                text = obj.text or ""
                if content_hash(text) == target_hash:
                    if occurrence_count == target_occurrence:
                        return block_type, obj, el, occurrence_count
                    occurrence_count += 1
    raise ValueError(f"Block not found: {target_id}")


def count_occurrence(
    doc: Document, block_type: str, text: str, target_el: CT_P | CT_Tbl
) -> int:
    """Count how many blocks with same type+hash appear before target_el.

    Used after edits to compute the correct occurrence index for returned ID.
    For tables, 'text' should be the full table content (from table_content_for_hash).
    """
    target_hash = content_hash(text)
    occurrence = 0
    for kind, obj, el in iter_body_blocks(doc):
        if kind == "table":
            if block_type == "table":
                table_content = table_content_for_hash(obj)
                if content_hash(table_content) == target_hash:
                    if el is target_el:
                        return occurrence
                    occurrence += 1
        else:
            actual_type, _ = paragraph_kind_and_level(obj)
            if block_type == actual_type:
                block_text = obj.text or ""
                if content_hash(block_text) == target_hash:
                    if el is target_el:
                        return occurrence
                    occurrence += 1
    raise ValueError("target_el not found in document")


def insert_paragraph_relative(
    doc: Document,
    target_el: CT_P | CT_Tbl,
    text: str,
    position: str,
    style_name: str = "",
) -> Paragraph:
    """Insert paragraph before/after target element."""
    new_p = doc.add_paragraph(text, style_name or None)
    if position == "before":
        target_el.addprevious(new_p._element)
    else:
        target_el.addnext(new_p._element)
    return new_p


def insert_heading_relative(
    doc: Document, target_el: CT_P | CT_Tbl, text: str, level: int, position: str
) -> Paragraph:
    """Insert heading before/after target element."""
    new_p = doc.add_heading(text, level=level)
    if position == "before":
        target_el.addprevious(new_p._element)
    else:
        target_el.addnext(new_p._element)
    return new_p


def insert_table_relative(
    doc: Document,
    target_el: CT_P | CT_Tbl,
    table_data: list[list[str]],
    position: str,
    style_name: str = "Table Grid",
) -> Table:
    """Insert table before/after target element."""
    rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.style = style_name
    for r in range(rows):
        for c in range(len(table_data[r])):
            tbl.cell(r, c).text = str(table_data[r][c])
    if position == "before":
        target_el.addprevious(tbl._tbl)
    else:
        target_el.addnext(tbl._tbl)
    return tbl


def delete_block(block_type: str, obj: Paragraph | Table) -> None:
    """Delete a block from the document."""
    # block_type is "paragraph", "heading1", "heading2", etc., or "table"
    el = (
        obj._element
        if block_type == "paragraph" or block_type.startswith("heading")
        else obj._tbl
    )
    el.getparent().remove(el)


def replace_table(doc: Document, old_tbl: Table, table_data: list[list[str]]) -> Table:
    """Replace table with new data."""
    rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
    new_tbl = doc.add_table(rows=rows, cols=cols)
    for r in range(rows):
        for c in range(len(table_data[r])):
            new_tbl.cell(r, c).text = str(table_data[r][c])
    old_el = old_tbl._tbl
    old_el.addprevious(new_tbl._tbl)
    old_el.getparent().remove(old_el)
    return new_tbl


def apply_paragraph_formatting(p: Paragraph, fmt: dict) -> None:
    """Apply direct formatting to paragraph (affects all runs)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    alignment_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if "alignment" in fmt:
        p.alignment = alignment_map[fmt["alignment"].lower()]

    for run in p.runs:
        for key, value in fmt.items():
            if key == "bold":
                run.bold = bool(value)
            elif key == "italic":
                run.italic = bool(value)
            elif key == "underline":
                run.underline = bool(value)
            elif key == "font_name":
                run.font.name = value
            elif key == "font_size":
                run.font.size = Pt(float(value))
            elif key == "color":
                run.font.color.rgb = RGBColor.from_string(value.lstrip("#"))


def build_table_cells(table: Table) -> list[CellInfo]:
    """Build list of CellInfo for all cells in a table."""
    return [
        CellInfo(row=r_idx + 1, col=c_idx + 1, text=cell.text or "")
        for r_idx, row in enumerate(table.rows)
        for c_idx, cell in enumerate(row.cells)
    ]


def replace_table_cell(table: Table, row: int, col: int, text: str) -> None:
    """Replace text in a table cell. Row/col are 1-based."""
    table.cell(row - 1, col - 1).text = text


def build_runs(paragraph: Paragraph) -> list[RunInfo]:
    """Build list of RunInfo for all runs in a paragraph."""
    return [
        RunInfo(
            index=idx,
            text=run.text or "",
            bold=run.bold,
            italic=run.italic,
            underline=run.underline,
            font_name=run.font.name,
            font_size=run.font.size.pt if run.font.size else None,
            color=str(run.font.color.rgb)
            if run.font.color and run.font.color.rgb
            else None,
        )
        for idx, run in enumerate(paragraph.runs)
    ]


def edit_run_text(paragraph: Paragraph, run_index: int, text: str) -> None:
    """Edit run text."""
    paragraph.runs[run_index].text = text


def edit_run_formatting(paragraph: Paragraph, run_index: int, fmt: dict) -> None:
    """Apply formatting to a specific run."""
    from docx.shared import Pt, RGBColor

    run = paragraph.runs[run_index]
    for key, value in fmt.items():
        if key == "bold":
            run.bold = bool(value)
        elif key == "italic":
            run.italic = bool(value)
        elif key == "underline":
            run.underline = bool(value)
        elif key == "font_name":
            run.font.name = value
        elif key == "font_size":
            run.font.size = Pt(float(value))
        elif key == "color":
            run.font.color.rgb = RGBColor.from_string(value.lstrip("#"))


def build_comments(doc: Document) -> list[CommentInfo]:
    """Build list of CommentInfo from document comments."""
    return [
        CommentInfo(
            id=c.comment_id,
            author=c.author,
            initials=c.initials,
            timestamp=c.timestamp.isoformat() if c.timestamp else None,
            text=c.text,
        )
        for c in doc.comments
    ]


def add_comment_to_block(
    doc: Document,
    paragraph: Paragraph,
    text: str,
    author: str = "",
    initials: str = "",
) -> int:
    """Add a comment anchored to all runs in a paragraph. Returns comment_id."""
    return doc.add_comment(
        runs=paragraph.runs, text=text, author=author, initials=initials
    ).comment_id


def build_headers_footers(doc: Document) -> list[HeaderFooterInfo]:
    """Build list of HeaderFooterInfo for all sections."""
    result = []
    for idx, section in enumerate(doc.sections):
        hdr, ftr = section.header, section.footer
        info = HeaderFooterInfo(
            section_index=idx,
            header_text=None
            if hdr.is_linked_to_previous
            else "\n".join(p.text for p in hdr.paragraphs),
            footer_text=None
            if ftr.is_linked_to_previous
            else "\n".join(p.text for p in ftr.paragraphs),
            header_is_linked=hdr.is_linked_to_previous,
            footer_is_linked=ftr.is_linked_to_previous,
            has_different_first_page=section.different_first_page_header_footer,
        )
        if section.different_first_page_header_footer:
            fp_hdr, fp_ftr = section.first_page_header, section.first_page_footer
            info.first_page_header_text = (
                None
                if fp_hdr.is_linked_to_previous
                else "\n".join(p.text for p in fp_hdr.paragraphs)
            )
            info.first_page_footer_text = (
                None
                if fp_ftr.is_linked_to_previous
                else "\n".join(p.text for p in fp_ftr.paragraphs)
            )
        result.append(info)
    return result


def set_header_text(doc: Document, section_index: int, text: str) -> None:
    """Set header text for a section."""
    header = doc.sections[section_index].header
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)
    header.add_paragraph(text)


def set_footer_text(doc: Document, section_index: int, text: str) -> None:
    """Set footer text for a section."""
    footer = doc.sections[section_index].footer
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)
    footer.add_paragraph(text)


def build_page_setup(doc: Document) -> list[PageSetupInfo]:
    """Build list of PageSetupInfo for all sections."""
    from docx.enum.section import WD_ORIENT

    result = []
    for idx, section in enumerate(doc.sections):
        s = section
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
            )
        )
    return result


def set_page_margins(
    doc: Document,
    section_index: int,
    top: float,
    bottom: float,
    left: float,
    right: float,
) -> None:
    """Set page margins for a section. Values in inches."""
    from docx.shared import Emu

    section = doc.sections[section_index]
    section.top_margin = Emu(int(top * 914400))
    section.bottom_margin = Emu(int(bottom * 914400))
    section.left_margin = Emu(int(left * 914400))
    section.right_margin = Emu(int(right * 914400))


def set_page_orientation(doc: Document, section_index: int, orientation: str) -> None:
    """Set page orientation for a section. 'portrait' or 'landscape'."""
    from docx.enum.section import WD_ORIENT
    from docx.shared import Emu

    section = doc.sections[section_index]
    w, h = section.page_width, section.page_height

    orient_lower = orientation.lower()
    section.orientation = (
        WD_ORIENT.LANDSCAPE if orient_lower == "landscape" else WD_ORIENT.PORTRAIT
    )
    if orient_lower == "landscape" and h > w or orient_lower == "portrait" and w > h:
        section.page_width, section.page_height = Emu(h), Emu(w)
