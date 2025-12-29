"""Python-docx wrapper functions for Word document manipulation."""

import hashlib
import json
import re
from collections.abc import Iterator
from dataclasses import dataclass

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_BREAK,
    WD_COLOR_INDEX,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap as oxml_nsmap
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl, CT_Tc
from docx.oxml.text.paragraph import CT_P
from docx.shared import Emu, Inches, Pt, RGBColor
from docx.table import Table, _Cell
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph

from mcp_handley_lab.word.models import (
    Block,
    CellInfo,
    CommentInfo,
    DocumentMeta,
    HeaderFooterInfo,
    HyperlinkInfo,
    ImageInfo,
    PageSetupInfo,
    ParagraphFormatInfo,
    RowInfo,
    RunInfo,
    StyleFormatInfo,
    StyleInfo,
    TableLayoutInfo,
    TabStopInfo,
)

_HEADING_RE = re.compile(r"^Heading ([1-9])$")
_ID_RE = re.compile(r"^(paragraph|heading[1-9]|table)_([0-9a-f]{8})_(\d+)$")
_IMAGE_ID_RE = re.compile(r"^image_([0-9a-f]{8})_(\d+)$")
_EMU_PER_INCH = 914400

# Highlight color mapping (WD_COLOR_INDEX enum)
_HIGHLIGHT_MAP = {
    "yellow": WD_COLOR_INDEX.YELLOW,
    "green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE,
    "pink": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE,
    "red": WD_COLOR_INDEX.RED,
    "dark_blue": WD_COLOR_INDEX.DARK_BLUE,
    "dark_red": WD_COLOR_INDEX.DARK_RED,
    "dark_yellow": WD_COLOR_INDEX.DARK_YELLOW,
    "gray": WD_COLOR_INDEX.GRAY_25,
    "dark_gray": WD_COLOR_INDEX.GRAY_50,
    "black": WD_COLOR_INDEX.BLACK,
    "white": WD_COLOR_INDEX.WHITE,
}
_HIGHLIGHT_REVERSE = {v: k for k, v in _HIGHLIGHT_MAP.items()}

# Section start type mapping
_SECTION_START_MAP = {
    "new_page": WD_SECTION.NEW_PAGE,
    "continuous": WD_SECTION.CONTINUOUS,
    "even_page": WD_SECTION.EVEN_PAGE,
    "odd_page": WD_SECTION.ODD_PAGE,
    "new_column": WD_SECTION.NEW_COLUMN,
}

# Hierarchical path segment patterns (0-based indices)
_CELL_RE = re.compile(r"^r(\d+)c(\d+)$")
_PARA_RE = re.compile(r"^p(\d+)$")
_TABLE_RE = re.compile(r"^tbl(\d+)$")

# Run formatting attribute mappings (attr_path, transform)
_RUN_ATTRS = {
    "bold": "bold",
    "italic": "italic",
    "underline": "underline",
    "strike": "font.strike",
    "double_strike": "font.double_strike",
    "subscript": "font.subscript",
    "superscript": "font.superscript",
    "all_caps": "font.all_caps",
    "small_caps": "font.small_caps",
    "hidden": "font.hidden",
    "emboss": "font.emboss",
    "imprint": "font.imprint",
    "outline": "font.outline",
    "shadow": "font.shadow",
    "font_name": "font.name",
}


def populate_table(table: Table, data: list[list]) -> None:
    """Populate table cells from 2D list."""
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            table.cell(r, c).text = str(val)


# PathSegment: (kind, indices) where kind='cell'|'para'|'tbl'
PathSegment = tuple[str, tuple[int, ...]]


@dataclass
class ResolvedTarget:
    """Result of resolving a hierarchical target ID."""

    base_id: str
    base_kind: str  # 'table' | 'paragraph' | 'heading1' etc
    base_obj: Paragraph | Table
    base_occurrence: int
    leaf_kind: str  # 'table' | 'cell' | 'paragraph'
    leaf_obj: Paragraph | Table | _Cell
    leaf_el: CT_P | CT_Tbl | CT_Tc  # XML element (Paragraph, Table, or Cell)


def content_hash(text: str) -> str:
    """8-char SHA256 of normalized text for content-addressable IDs."""
    normalized = re.sub(r"\s+", " ", text.strip().lower())
    return hashlib.sha256(normalized.encode()).hexdigest()[:8]


def make_block_id(block_type: str, text: str, occurrence: int) -> str:
    """Generate content-addressable block ID. Single source of truth."""
    return f"{block_type}_{content_hash(text)}_{occurrence}"


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


def create_element(
    doc: Document,
    content_type: str,
    content_data: str,
    style_name: str = "",
    heading_level: int = 0,
) -> Paragraph | Table:
    """Create a content element (paragraph, heading, or table)."""
    if content_type == "paragraph":
        return doc.add_paragraph(content_data, style_name or None)
    if content_type == "heading":
        return doc.add_heading(content_data, level=heading_level)
    if content_type == "table":
        table_data = json.loads(content_data)
        rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
        tbl = doc.add_table(rows=rows, cols=cols)
        tbl.style = style_name or "Table Grid"
        populate_table(tbl, table_data)
        return tbl
    raise ValueError(f"Unknown content_type: {content_type}")


def get_element_id(
    doc: Document, obj: Paragraph | Table, heading_level: int = 0
) -> str:
    """Calculate content-addressed ID for an element."""
    if isinstance(obj, Table):
        block_type = "table"
        content = table_content_for_hash(obj)
        el = obj._tbl
    else:
        block_type, _ = paragraph_kind_and_level(obj)
        # Override block_type if heading_level specified (for newly created headings)
        if heading_level:
            block_type = f"heading{heading_level}"
        content = obj.text or ""
        el = obj._element
    occurrence = count_occurrence(doc, block_type, content, el)
    return make_block_id(block_type, content, occurrence)


def insert_content_relative(
    doc: Document,
    target_el: CT_P | CT_Tbl,
    position: str,
    content_type: str,
    content_data: str,
    style_name: str = "",
    heading_level: int = 0,
) -> Paragraph | Table:
    """Insert content relative to a target element."""
    obj = create_element(doc, content_type, content_data, style_name, heading_level)
    el = obj._tbl if isinstance(obj, Table) else obj._element
    _insert_at(target_el, el, position)
    return obj


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


def set_document_meta(doc: Document, **kwargs) -> None:
    """Update document core properties. Only updates non-None values."""
    cp = doc.core_properties
    for key, value in kwargs.items():
        if value is not None:
            setattr(cp, key, value)


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
                        id=make_block_id(block_type, text, occurrence),
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
                        id=make_block_id("table", table_content, occurrence),
                        type="table",
                        text=md,  # Keep markdown for display
                        style=style,
                        rows=rows,
                        cols=cols,
                    )
                )
            matched += 1

    return blocks, matched


def parse_target_id(target_id: str) -> tuple[str, list[PathSegment]]:
    """Parse target_id into (base_id, path_segments).

    Supports hierarchical IDs: table_abc12345_0#r0c1/p0
    Returns (base_id, []) for simple IDs without path.
    """
    base_id, *rest = target_id.split("#", 1)
    if not rest:
        return target_id, []

    segments: list[PathSegment] = []
    for part in rest[0].split("/"):
        if m := _CELL_RE.match(part):
            segments.append(("cell", (int(m[1]), int(m[2]))))
        elif m := _PARA_RE.match(part):
            segments.append(("para", (int(m[1]),)))
        elif m := _TABLE_RE.match(part):
            segments.append(("tbl", (int(m[1]),)))
        else:
            raise ValueError(f"Invalid path segment: {part}")
    return base_id, segments


def resolve_path(
    container: Table | _Cell, segments: list[PathSegment]
) -> Paragraph | Table | _Cell:
    """Resolve path segments from container. Returns leaf object.

    Transition rules:
    - From Table -> cell (r,c) only
    - From Cell -> para (p) or tbl (tbl)
    - From nested Table -> cell (r,c)
    """
    current: Table | _Cell | Paragraph = container

    for kind, indices in segments:
        if kind == "cell":
            row, col = indices
            current = current.cell(
                row, col
            )  # AttributeError if not Table, IndexError if OOB
        elif kind == "para":
            current = current.paragraphs[indices[0]]  # AttributeError if not Cell
        elif kind == "tbl":
            current = current.tables[indices[0]]  # AttributeError if not Cell

    return current


def _resolve_base_block(
    doc: Document, base_id: str
) -> tuple[str, Paragraph | Table, CT_P | CT_Tbl, int]:
    """Resolve base block ID to (block_type, obj, element, occurrence).

    Uses content-hash IDs: {type}_{hash}_{occurrence}
    Searches all blocks for matching type+hash, then skips to Nth occurrence.
    """
    target_type, target_hash, occurrence_str = _ID_RE.match(base_id).groups()
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
    raise ValueError(f"Block not found: {base_id}")


def resolve_target(doc: Document, target_id: str) -> ResolvedTarget:
    """Resolve target_id to ResolvedTarget with base and leaf info.

    Supports hierarchical IDs: table_abc12345_0#r0c1/p0
    """
    base_id, path_segments = parse_target_id(target_id)

    # Resolve base block
    base_kind, base_obj, base_el, base_occurrence = _resolve_base_block(doc, base_id)

    if not path_segments:
        return ResolvedTarget(
            base_id=base_id,
            base_kind=base_kind,
            base_obj=base_obj,
            base_occurrence=base_occurrence,
            leaf_kind=base_kind,
            leaf_obj=base_obj,
            leaf_el=base_el,
        )

    # Resolve path within container (AttributeError if base_obj not a Table)
    leaf_obj = resolve_path(base_obj, path_segments)

    # Determine leaf kind and element based on type
    if isinstance(leaf_obj, Table):
        return ResolvedTarget(
            base_id=base_id,
            base_kind=base_kind,
            base_obj=base_obj,
            base_occurrence=base_occurrence,
            leaf_kind="table",
            leaf_obj=leaf_obj,
            leaf_el=leaf_obj._tbl,
        )
    if isinstance(leaf_obj, _Cell):
        return ResolvedTarget(
            base_id=base_id,
            base_kind=base_kind,
            base_obj=base_obj,
            base_occurrence=base_occurrence,
            leaf_kind="cell",
            leaf_obj=leaf_obj,
            leaf_el=leaf_obj._tc,
        )
    # Paragraph
    return ResolvedTarget(
        base_id=base_id,
        base_kind=base_kind,
        base_obj=base_obj,
        base_occurrence=base_occurrence,
        leaf_kind="paragraph",
        leaf_obj=leaf_obj,
        leaf_el=leaf_obj._element,
    )


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


def _insert_at(target_el, new_el, position: str) -> None:
    """Insert new_el before or after target_el."""
    (target_el.addprevious if position == "before" else target_el.addnext)(new_el)


def insert_paragraph_relative(
    doc: Document,
    target_el: CT_P | CT_Tbl,
    text: str,
    position: str,
    style_name: str = "",
) -> Paragraph:
    """Insert paragraph before/after target element."""
    new_p = doc.add_paragraph(text, style_name or None)
    _insert_at(target_el, new_p._element, position)
    return new_p


def insert_heading_relative(
    doc: Document, target_el: CT_P | CT_Tbl, text: str, level: int, position: str
) -> Paragraph:
    """Insert heading before/after target element."""
    new_p = doc.add_heading(text, level=level)
    _insert_at(target_el, new_p._element, position)
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
    populate_table(tbl, table_data)
    _insert_at(target_el, tbl._tbl, position)
    return tbl


def add_page_break(doc: Document) -> Paragraph:
    """Append a page break paragraph to the document. Returns the paragraph."""

    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)
    return p


def add_break_after(
    doc: Document, target_el: CT_P | CT_Tbl, break_type: str
) -> Paragraph:
    """Insert break after target element. break_type: 'page', 'column', 'line'."""

    break_map = {
        "page": WD_BREAK.PAGE,
        "column": WD_BREAK.COLUMN,
        "line": WD_BREAK.LINE,
    }
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(break_map[break_type])  # KeyError if invalid
    target_el.addnext(p._element)
    return p


def delete_block(obj: Paragraph | Table) -> None:
    """Delete a block from the document."""
    el = obj._element if isinstance(obj, Paragraph) else obj._tbl
    el.getparent().remove(el)


def replace_table(doc: Document, old_tbl: Table, table_data: list[list[str]]) -> Table:
    """Replace table with new data."""
    rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
    new_tbl = doc.add_table(rows=rows, cols=cols)
    populate_table(new_tbl, table_data)
    old_el = old_tbl._tbl
    old_el.addprevious(new_tbl._tbl)
    old_el.getparent().remove(old_el)
    return new_tbl


_PARA_INCH_ATTRS = {"left_indent", "right_indent", "first_line_indent"}
_PARA_PT_ATTRS = {"space_before", "space_after"}
_PARA_DIRECT_ATTRS = {"keep_with_next", "page_break_before"}
_RUN_FORMAT_KEYS = set(_RUN_ATTRS) | {"style", "font_size", "color", "highlight_color"}


def apply_paragraph_formatting(p: Paragraph, fmt: dict) -> None:
    """Apply direct formatting to paragraph (affects all runs)."""
    if "alignment" in fmt:
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, fmt["alignment"].upper())

    pf = p.paragraph_format
    for key, value in fmt.items():
        if key in _PARA_INCH_ATTRS:
            setattr(pf, key, Inches(value))
        elif key in _PARA_PT_ATTRS:
            setattr(pf, key, Pt(value))
        elif key == "line_spacing":
            pf.line_spacing = value if value < 5 else Pt(value)
        elif key in _PARA_DIRECT_ATTRS:
            setattr(pf, key, value)

    for run in p.runs:
        for key, value in fmt.items():
            if key in _RUN_FORMAT_KEYS:
                _set_run_attr(run, key, value)


def _get_vmerge_val_from_tc(tc: CT_Tc) -> str | None:
    """Get vMerge value from a tc element.

    Returns:
        'restart' for merge origin, 'continue' for continuation, None for no merge.
    """

    tc_pr = tc.find(qn("w:tcPr"))
    if tc_pr is None:
        return None
    v_merge = tc_pr.find(qn("w:vMerge"))
    if v_merge is None:
        return None
    # vMerge with val="restart" is origin, vMerge without val is continue
    val = v_merge.get(qn("w:val"))
    return val if val else "continue"


def _tc_at_grid_col(tr, grid_col: int):
    """Find the tc element at a given grid column, accounting for gridSpan.

    Returns the tc element or None if not found.
    """

    c = 0
    for tc in tr.findall(qn("w:tc")):
        span = tc.grid_span
        if c == grid_col:
            return tc
        c += span
        if c > grid_col:
            # Grid column falls within a span, no tc origin at this position
            return None
    return None


def _calculate_row_span_from_xml(table: Table, start_row: int, col: int) -> int:
    """Calculate vertical span by checking vMerge='continue' in subsequent rows.

    Iterates over actual XML tc elements to detect continuation cells.
    """
    span = 1
    rows = list(table.rows)
    for r in range(start_row + 1, len(rows)):
        row_el = rows[r]._tr
        tc = _tc_at_grid_col(row_el, col)
        if tc is not None:
            vmerge = _get_vmerge_val_from_tc(tc)
            if vmerge == "continue":
                span += 1
            else:
                break
        else:
            break
    return span


def build_table_cells(table: Table, table_id: str = "") -> list[CellInfo]:
    """Build list of CellInfo with merge information.

    Detects:
    - Horizontal merges via grid_span property
    - Vertical merges via vMerge XML attribute

    Iterates over actual XML elements (not table.cell()) to correctly
    detect continuation cells in vertical merges.

    Args:
        table: The Table object
        table_id: Base ID of the table (for hierarchical IDs)

    Returns:
        List of CellInfo with merge info for all grid positions
    """

    valign_map = {
        WD_CELL_VERTICAL_ALIGNMENT.TOP: "top",
        WD_CELL_VERTICAL_ALIGNMENT.CENTER: "center",
        WD_CELL_VERTICAL_ALIGNMENT.BOTTOM: "bottom",
    }

    result = []
    rows = list(table.rows)

    for r, row in enumerate(rows):
        row_el = row._tr
        tc_elements = row_el.findall(qn("w:tc"))
        c = 0
        for tc in tc_elements:
            vmerge = _get_vmerge_val_from_tc(tc)
            grid_span = tc.grid_span

            if vmerge == "continue":
                # Vertical continuation cell
                result.append(
                    CellInfo(
                        row=r,
                        col=c,
                        text="",
                        hierarchical_id=f"{table_id}#r{r}c{c}" if table_id else "",
                        is_merge_origin=False,
                        grid_span=1,
                        row_span=1,
                    )
                )
                # Add horizontal continuation entries for wide continuation cells
                for span_c in range(1, grid_span):
                    result.append(
                        CellInfo(
                            row=r,
                            col=c + span_c,
                            text="",
                            hierarchical_id=(
                                f"{table_id}#r{r}c{c + span_c}" if table_id else ""
                            ),
                            is_merge_origin=False,
                            grid_span=1,
                            row_span=1,
                        )
                    )
                c += grid_span
            else:
                # Origin cell (vmerge='restart' or None) or normal cell
                row_span = (
                    _calculate_row_span_from_xml(table, r, c)
                    if vmerge == "restart"
                    else 1
                )
                # Get text and properties from the cell
                cell = table.cell(r, c)
                width_inches = cell.width.inches if cell.width else None
                valign = valign_map.get(cell.vertical_alignment)
                result.append(
                    CellInfo(
                        row=r,
                        col=c,
                        text=cell.text or "",
                        hierarchical_id=f"{table_id}#r{r}c{c}" if table_id else "",
                        is_merge_origin=True,
                        grid_span=grid_span,
                        row_span=row_span,
                        width_inches=width_inches,
                        vertical_alignment=valign,
                    )
                )
                # Add continuation entries for horizontal span
                for span_c in range(1, grid_span):
                    result.append(
                        CellInfo(
                            row=r,
                            col=c + span_c,
                            text="",
                            hierarchical_id=(
                                f"{table_id}#r{r}c{c + span_c}" if table_id else ""
                            ),
                            is_merge_origin=False,
                            grid_span=1,
                            row_span=1,
                        )
                    )
                c += grid_span
    return result


def merge_cells(
    table: Table, start_row: int, start_col: int, end_row: int, end_col: int
) -> None:
    """Merge a rectangular region of cells."""
    start_cell = table.cell(start_row, start_col)
    end_cell = table.cell(end_row, end_col)
    start_cell.merge(end_cell)


def replace_table_cell(table: Table, row: int, col: int, text: str) -> None:
    """Replace text in a table cell. Row/col are 0-based."""
    table.cell(row, col).text = text


def add_table_row(table: Table, data: list[str] | None = None) -> int:
    """Add row to table. Returns new row index (0-based)."""
    row = table.add_row()
    if data:
        for i, text in enumerate(data[: len(table.columns)]):
            row.cells[i].text = text
    return len(table.rows) - 1


def add_table_column(
    table: Table, width_inches: float = 1.0, data: list[str] | None = None
) -> int:
    """Add column to table. Width required by python-docx API. Returns new col index."""

    table.add_column(Inches(width_inches))
    col_idx = len(table.columns) - 1
    if data:
        for i, text in enumerate(data[: len(table.rows)]):
            table.cell(i, col_idx).text = text
    return col_idx


def delete_table_row(table: Table, row_index: int) -> None:
    """Delete row from table (0-based index)."""
    row = table.rows[row_index]
    row._element.getparent().remove(row._element)


def delete_table_column(table: Table, col_index: int) -> None:
    """Delete column from table (0-based index). Removes grid definition and cells."""
    # 1. Remove the grid column definition (required for valid Word XML)
    tbl_grid = table._tbl.tblGrid
    if col_index < len(tbl_grid.gridCol_lst):
        grid_col = tbl_grid.gridCol_lst[col_index]
        grid_col.getparent().remove(grid_col)

    # 2. Remove the cell from every row
    for row in table.rows:
        if col_index < len(row.cells):
            cell = row.cells[col_index]
            cell._element.getparent().remove(cell._element)


# --- Table Layout Functions ---


def build_table_layout(table: Table, table_id: str) -> TableLayoutInfo:
    """Build table layout info including row heights and alignment."""

    table_align_map = {
        WD_TABLE_ALIGNMENT.LEFT: "left",
        WD_TABLE_ALIGNMENT.CENTER: "center",
        WD_TABLE_ALIGNMENT.RIGHT: "right",
    }
    row_height_rule_map = {
        WD_ROW_HEIGHT_RULE.AUTO: "auto",
        WD_ROW_HEIGHT_RULE.AT_LEAST: "at_least",
        WD_ROW_HEIGHT_RULE.EXACTLY: "exactly",
    }

    rows = []
    for i, row in enumerate(table.rows):
        rows.append(
            RowInfo(
                index=i,
                height_inches=row.height.inches if row.height else None,
                height_rule=row_height_rule_map.get(row.height_rule),
            )
        )

    return TableLayoutInfo(
        table_id=table_id,
        alignment=table_align_map.get(table.alignment),
        autofit=table.autofit,
        rows=rows,
    )


def set_table_alignment(table: Table, alignment: str) -> None:
    """Set table horizontal alignment."""

    alignment_map = {
        "left": WD_TABLE_ALIGNMENT.LEFT,
        "center": WD_TABLE_ALIGNMENT.CENTER,
        "right": WD_TABLE_ALIGNMENT.RIGHT,
    }
    table.alignment = alignment_map[alignment.lower()]  # KeyError if invalid


def set_row_height(
    table: Table, row_index: int, height_inches: float, rule: str = "at_least"
) -> None:
    """Set row height. Default rule is 'at_least' to prevent text clipping."""

    rule_map = {
        "auto": WD_ROW_HEIGHT_RULE.AUTO,
        "at_least": WD_ROW_HEIGHT_RULE.AT_LEAST,
        "exactly": WD_ROW_HEIGHT_RULE.EXACTLY,
    }
    rule_val = rule.lower()
    row = table.rows[row_index]  # Let IndexError propagate
    row.height = None if rule_val == "auto" else Inches(height_inches)
    row.height_rule = rule_map[rule_val]  # KeyError if invalid


def set_table_fixed_layout(table: Table, column_widths: list[float]) -> None:
    """Set table to fixed layout with explicit column widths (inches)."""

    table.autofit = False
    for i, width in enumerate(column_widths):
        if i < len(table.columns):
            table.columns[i].width = Inches(width)


def set_cell_width(table: Table, row: int, col: int, width_inches: float) -> None:
    """Set cell width."""

    table.cell(row, col).width = Inches(width_inches)


def set_cell_vertical_alignment(
    table: Table, row: int, col: int, alignment: str
) -> None:
    """Set cell vertical alignment."""

    valign_map = {
        "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
        "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
        "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
    }
    table.cell(row, col).vertical_alignment = valign_map[alignment.lower()]


def _build_run_info(
    run, index: int, is_hyperlink: bool = False, hyperlink_url: str | None = None
) -> RunInfo:
    """Build RunInfo from a Run object."""
    return RunInfo(
        index=index,
        text=run.text or "",
        bold=run.bold,
        italic=run.italic,
        underline=run.underline,
        font_name=run.font.name,
        font_size=run.font.size.pt if run.font.size else None,
        color=str(run.font.color.rgb)
        if run.font.color and run.font.color.rgb
        else None,
        highlight_color=_HIGHLIGHT_REVERSE.get(run.font.highlight_color),
        strike=run.font.strike,
        double_strike=run.font.double_strike,
        subscript=run.font.subscript,
        superscript=run.font.superscript,
        style=run.style.name if run.style else None,
        is_hyperlink=is_hyperlink,
        hyperlink_url=hyperlink_url,
        # Additional font properties
        all_caps=run.font.all_caps,
        small_caps=run.font.small_caps,
        hidden=run.font.hidden,
        emboss=run.font.emboss,
        imprint=run.font.imprint,
        outline=run.font.outline,
        shadow=run.font.shadow,
    )


def build_runs(paragraph: Paragraph) -> list[RunInfo]:
    """Build list of RunInfo for all runs in a paragraph, including hyperlink runs."""

    result = []
    idx = 0
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            url = item.url
            for run in item.runs:
                result.append(
                    _build_run_info(run, idx, is_hyperlink=True, hyperlink_url=url)
                )
                idx += 1
        else:  # Run
            result.append(_build_run_info(item, idx))
            idx += 1
    return result


def build_hyperlinks(doc: Document) -> list[HyperlinkInfo]:
    """Build list of all hyperlinks in the document."""

    result = []
    idx = 0
    for para, _el in _iter_all_paragraphs(doc):
        for item in para.iter_inner_content():
            if isinstance(item, Hyperlink):
                result.append(
                    HyperlinkInfo(
                        index=idx,
                        text=item.text,
                        url=item.url,
                        address=item.address,
                        fragment=item.fragment,
                        is_external=bool(item._hyperlink.rId),
                    )
                )
                idx += 1
    return result


def build_styles(doc: Document) -> list[StyleInfo]:
    """Build list of all styles in the document."""

    style_type_map = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    result = []
    for style in doc.styles:
        # Some style types (e.g., _NumberingStyle) don't have all attributes
        base = getattr(style, "base_style", None)
        next_style = getattr(style, "next_paragraph_style", None)
        hidden = getattr(style, "hidden", False)
        quick_style = getattr(style, "quick_style", False)
        result.append(
            StyleInfo(
                name=style.name,
                style_id=style.style_id,
                type=style_type_map.get(style.type, "unknown"),
                builtin=style.builtin,
                base_style=base.name if base else None,
                next_style=(
                    next_style.name if next_style and next_style != style else None
                ),
                hidden=hidden,
                quick_style=quick_style,
            )
        )
    return result


def get_style_format(doc: Document, style_name: str) -> StyleFormatInfo:
    """Get detailed formatting for a specific style."""

    # Alignment normalization map (enum -> API value)
    alignment_to_api = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }

    style = doc.styles[style_name]
    font = style.font

    # Determine style type
    style_type_map = {
        WD_STYLE_TYPE.PARAGRAPH: "paragraph",
        WD_STYLE_TYPE.CHARACTER: "character",
        WD_STYLE_TYPE.TABLE: "table",
        WD_STYLE_TYPE.LIST: "list",
    }
    style_type = style_type_map.get(style.type, "unknown")

    # Extract paragraph format if available
    pf = getattr(style, "paragraph_format", None)
    alignment = None
    left_indent = None
    space_before = None
    space_after = None
    line_spacing = None

    if pf:
        alignment = alignment_to_api.get(pf.alignment) if pf.alignment else None
        left_indent = pf.left_indent.inches if pf.left_indent else None
        space_before = pf.space_before.pt if pf.space_before else None
        space_after = pf.space_after.pt if pf.space_after else None
        # Handle both float multipliers and Length objects (same as build_paragraph_format)
        line_spacing = (
            pf.line_spacing
            if isinstance(pf.line_spacing, float)
            else (pf.line_spacing.pt if pf.line_spacing else None)
        )

    return StyleFormatInfo(
        name=style.name,
        style_id=style.style_id,
        type=style_type,
        font_name=font.name,
        font_size=font.size.pt if font.size else None,
        bold=font.bold,
        italic=font.italic,
        color=str(font.color.rgb) if font.color and font.color.rgb else None,
        alignment=alignment,
        left_indent=left_indent,
        space_before=space_before,
        space_after=space_after,
        line_spacing=line_spacing,
    )


def edit_style(doc: Document, style_name: str, fmt: dict) -> None:
    """Modify a style definition."""

    style = doc.styles[style_name]
    font = style.font

    # Font properties
    if "font_name" in fmt:
        font.name = fmt["font_name"]
    if "font_size" in fmt:
        font.size = Pt(fmt["font_size"])
    if "bold" in fmt:
        font.bold = fmt["bold"]
    if "italic" in fmt:
        font.italic = fmt["italic"]
    if "color" in fmt:
        font.color.rgb = RGBColor.from_string(fmt["color"].lstrip("#"))

    # Paragraph properties (paragraph/table styles only)
    pf = getattr(style, "paragraph_format", None)
    if pf:
        if "alignment" in fmt:
            pf.alignment = getattr(WD_ALIGN_PARAGRAPH, fmt["alignment"].upper())
        if "left_indent" in fmt:
            pf.left_indent = Inches(fmt["left_indent"])
        if "space_before" in fmt:
            pf.space_before = Pt(fmt["space_before"])
        if "space_after" in fmt:
            pf.space_after = Pt(fmt["space_after"])
        if "line_spacing" in fmt:
            val = fmt["line_spacing"]
            pf.line_spacing = val if val < 5 else Pt(val)


def build_tab_stops(paragraph: Paragraph) -> list[TabStopInfo]:
    """Build list of tab stops for a paragraph."""

    tab_align_map = {
        WD_TAB_ALIGNMENT.LEFT: "left",
        WD_TAB_ALIGNMENT.CENTER: "center",
        WD_TAB_ALIGNMENT.RIGHT: "right",
        WD_TAB_ALIGNMENT.DECIMAL: "decimal",
        WD_TAB_ALIGNMENT.BAR: "bar",
    }
    tab_leader_map = {
        WD_TAB_LEADER.SPACES: "spaces",
        WD_TAB_LEADER.DOTS: "dots",
        WD_TAB_LEADER.HEAVY: "heavy",
        WD_TAB_LEADER.MIDDLE_DOT: "middle_dot",
        None: "spaces",  # None means default (spaces)
    }

    result = []
    for tab in paragraph.paragraph_format.tab_stops:
        alignment = tab_align_map.get(tab.alignment, "unknown")
        leader = tab_leader_map.get(tab.leader, "unknown")
        result.append(
            TabStopInfo(
                position_inches=round(tab.position.inches, 4),
                alignment=alignment,
                leader=leader,
            )
        )
    return result


def add_tab_stop(
    paragraph: Paragraph,
    position_inches: float,
    alignment: str = "left",
    leader: str = "spaces",
) -> None:
    """Add a tab stop to a paragraph."""

    align_map = {
        "left": WD_TAB_ALIGNMENT.LEFT,
        "center": WD_TAB_ALIGNMENT.CENTER,
        "right": WD_TAB_ALIGNMENT.RIGHT,
        "decimal": WD_TAB_ALIGNMENT.DECIMAL,
    }
    leader_map = {
        "spaces": WD_TAB_LEADER.SPACES,
        "dots": WD_TAB_LEADER.DOTS,
        "heavy": WD_TAB_LEADER.HEAVY,
        "middle_dot": WD_TAB_LEADER.MIDDLE_DOT,
    }
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(position_inches),
        align_map[alignment.lower()],  # KeyError if invalid
        leader_map[leader.lower()],  # KeyError if invalid
    )


# --- Document Fields ---


def _make_run_with(*elements):
    """Create a w:r element containing the given child elements."""
    r = OxmlElement("w:r")
    for elem in elements:
        r.append(elem)
    return r


def insert_field(paragraph: Paragraph, field_code: str, display_text: str = "") -> None:
    """Insert a Word field into a paragraph.

    Creates proper OXML field structure with separate runs for each part:
    begin, instruction, separator, result, and end markers.
    Supports any Word field code (PAGE, NUMPAGES, DATE, TIME, AUTHOR, etc.).
    """
    code_upper = field_code.strip().upper()
    p = paragraph._p

    # Run 1: Field begin
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    p.append(_make_run_with(fld_char_begin))

    # Run 2: Field instruction
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = f" {code_upper} "
    p.append(_make_run_with(instr_text))

    # Run 3: Field separator
    fld_char_sep = OxmlElement("w:fldChar")
    fld_char_sep.set(qn("w:fldCharType"), "separate")
    p.append(_make_run_with(fld_char_sep))

    # Run 4: Result text (placeholder shown before field updates)
    text_elem = OxmlElement("w:t")
    text_elem.text = display_text or "1"
    p.append(_make_run_with(text_elem))

    # Run 5: Field end
    fld_char_end = OxmlElement("w:fldChar")
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


def build_paragraph_format(paragraph: Paragraph) -> ParagraphFormatInfo:
    """Extract paragraph formatting properties."""

    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    pf = paragraph.paragraph_format
    alignment = alignment_map.get(paragraph.alignment)
    return ParagraphFormatInfo(
        alignment=alignment,
        left_indent=pf.left_indent.inches if pf.left_indent else None,
        right_indent=pf.right_indent.inches if pf.right_indent else None,
        first_line_indent=pf.first_line_indent.inches if pf.first_line_indent else None,
        space_before=pf.space_before.pt if pf.space_before else None,
        space_after=pf.space_after.pt if pf.space_after else None,
        line_spacing=(
            pf.line_spacing
            if isinstance(pf.line_spacing, float)
            else (pf.line_spacing.pt if pf.line_spacing else None)
        ),
        keep_with_next=pf.keep_with_next,
        page_break_before=pf.page_break_before,
        tab_stops=build_tab_stops(paragraph),
    )


def _resolve_run_by_inner_index(paragraph: Paragraph, run_index: int):
    """Resolve run by iter_inner_content index (matching build_runs indexing).

    This ensures edit_run operations use the same indexing as read(scope='runs'),
    which includes runs inside hyperlinks.
    """

    idx = 0
    for item in paragraph.iter_inner_content():
        if isinstance(item, Hyperlink):
            for run in item.runs:
                if idx == run_index:
                    return run
                idx += 1
        else:  # Run
            if idx == run_index:
                return item
            idx += 1
    raise IndexError(f"Run index {run_index} out of range (paragraph has {idx} runs)")


def edit_run_text(paragraph: Paragraph, run_index: int, text: str) -> None:
    """Edit run text. Uses iter_inner_content indexing (includes hyperlink runs)."""
    run = _resolve_run_by_inner_index(paragraph, run_index)
    run.text = text


def _set_run_attr(run, key: str, value) -> None:
    """Set a run attribute by key. Handles nested paths like 'font.bold'."""
    if key == "style":
        run.style = value
    elif key == "font_size":
        run.font.size = Pt(float(value))
    elif key == "color":
        run.font.color.rgb = RGBColor.from_string(value.lstrip("#"))
    elif key == "highlight_color":
        run.font.highlight_color = _HIGHLIGHT_MAP[value.lower()]
    elif key in _RUN_ATTRS:
        path = _RUN_ATTRS[key]
        obj, attr = (run, path) if "." not in path else (run.font, path.split(".")[1])
        setattr(obj, attr, value)


def edit_run_formatting(paragraph: Paragraph, run_index: int, fmt: dict) -> None:
    """Apply formatting to a specific run. Uses iter_inner_content indexing."""
    run = _resolve_run_by_inner_index(paragraph, run_index)
    for key, value in fmt.items():
        _set_run_attr(run, key, value)


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
            has_different_odd_even=doc.settings.odd_and_even_pages_header_footer,
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
        if doc.settings.odd_and_even_pages_header_footer:
            ev_hdr, ev_ftr = section.even_page_header, section.even_page_footer
            info.even_page_header_text = (
                None
                if ev_hdr.is_linked_to_previous
                else "\n".join(p.text for p in ev_hdr.paragraphs)
            )
            info.even_page_footer_text = (
                None
                if ev_ftr.is_linked_to_previous
                else "\n".join(p.text for p in ev_ftr.paragraphs)
            )
        result.append(info)
    return result


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
        populate_table(tbl, table_data)
        return make_block_id("table", table_content_for_hash(tbl), 0)
    raise ValueError(f"Unknown content_type: {content_type}")


def clear_header_footer(doc: Document, section_index: int, location: str) -> None:
    """Clear header/footer content. Unlinks from previous section first."""
    hf = getattr(doc.sections[section_index], location)
    hf.is_linked_to_previous = False
    hf_el = hf._element
    for child in list(hf_el):
        hf_el.remove(child)
    hf.add_paragraph("")  # Word requires at least one paragraph


def build_page_setup(doc: Document) -> list[PageSetupInfo]:
    """Build list of PageSetupInfo for all sections."""

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

    section = doc.sections[section_index]
    section.top_margin = Emu(int(top * 914400))
    section.bottom_margin = Emu(int(bottom * 914400))
    section.left_margin = Emu(int(left * 914400))
    section.right_margin = Emu(int(right * 914400))


def set_page_orientation(doc: Document, section_index: int, orientation: str) -> None:
    """Set page orientation for a section. 'portrait' or 'landscape'."""

    section = doc.sections[section_index]
    w, h = section.page_width, section.page_height

    orient_lower = orientation.lower()
    section.orientation = (
        WD_ORIENT.LANDSCAPE if orient_lower == "landscape" else WD_ORIENT.PORTRAIT
    )
    if orient_lower == "landscape" and h > w or orient_lower == "portrait" and w > h:
        section.page_width, section.page_height = Emu(h), Emu(w)


def add_section(doc: Document, start_type: str = "new_page") -> int:
    """Add new section. Returns new section index (0-based).

    start_type: 'new_page', 'continuous', 'even_page', 'odd_page', 'new_column'.
    """
    section_start = _SECTION_START_MAP.get(start_type.lower(), WD_SECTION.NEW_PAGE)
    doc.add_section(section_start)
    return len(doc.sections) - 1


# Image support functions


def get_embedded_image_hash(doc: Document, blip) -> str | None:
    """Get SHA1 hash for embedded image. Returns None for linked images."""
    rel_id = blip.embed  # Only embedded, not linked
    if not rel_id:
        return None  # Linked image - can't hash external resource
    image_part = doc.part.related_parts[rel_id]
    return image_part.image.sha1[:8]


def _extract_images_from_run(
    doc: Document,
    run,
    run_idx: int,
    block_id: str,
    image_hash_counts: dict[str, int],
    images: list[ImageInfo],
) -> None:
    """Extract images from a single run."""
    image_idx_in_run = 0
    for drawing in run._element.findall(".//w:drawing", namespaces=oxml_nsmap):
        inline = drawing.find(".//wp:inline", namespaces=oxml_nsmap)
        if inline is None:
            continue

        # Guard access to pic element (skip charts/smartart)
        try:
            blip = inline.graphic.graphicData.pic.blipFill.blip
        except AttributeError:
            continue

        h = get_embedded_image_hash(doc, blip)
        if h is None:
            continue  # Skip linked images

        # Track image occurrence globally
        img_occurrence = image_hash_counts.get(h, 0)
        image_hash_counts[h] = img_occurrence + 1

        # Get metadata via XML (avoid InlineShape construction issues)
        image_part = doc.part.related_parts[blip.embed]
        extent = inline.extent
        width_emu = extent.cx if extent is not None else 0
        height_emu = extent.cy if extent is not None else 0

        images.append(
            ImageInfo(
                id=f"image_{h}_{img_occurrence}",
                width_inches=width_emu / _EMU_PER_INCH,
                height_inches=height_emu / _EMU_PER_INCH,
                content_type=image_part.content_type,
                block_id=block_id,
                run_index=run_idx,
                image_index_in_run=image_idx_in_run,
                filename=image_part.image.filename or "",
            )
        )
        image_idx_in_run += 1


def _extract_images_from_paragraph(
    doc: Document,
    para,
    block_id: str,
    image_hash_counts: dict[str, int],
) -> list[ImageInfo]:
    """Extract images from a paragraph, updating occurrence counts.

    Uses iter_inner_content() indexing to match build_runs() indexing.
    """

    images: list[ImageInfo] = []
    run_idx = 0
    for item in para.iter_inner_content():
        if isinstance(item, Hyperlink):
            for run in item.runs:
                _extract_images_from_run(
                    doc, run, run_idx, block_id, image_hash_counts, images
                )
                run_idx += 1
        else:  # Run
            _extract_images_from_run(
                doc, item, run_idx, block_id, image_hash_counts, images
            )
            run_idx += 1
    return images


def build_images(doc: Document) -> list[ImageInfo]:
    """Build list of ImageInfo from document body, including tables."""
    images = []
    block_hash_counts: dict[str, int] = {}  # For block_id computation
    image_hash_counts: dict[str, int] = {}  # For image occurrence

    for kind, obj, _el in iter_body_blocks(doc):
        if kind == "paragraph":
            # Use SAME logic as build_blocks for block_id
            block_type, _ = paragraph_kind_and_level(obj)
            text = obj.text or ""
            block_hash_key = f"{block_type}_{content_hash(text)}"
            block_occurrence = block_hash_counts.get(block_hash_key, 0)
            block_id = make_block_id(block_type, text, block_occurrence)
            block_hash_counts[block_hash_key] = block_occurrence + 1

            images.extend(
                _extract_images_from_paragraph(doc, obj, block_id, image_hash_counts)
            )

        elif kind == "table":
            # Compute table's block_id
            table_content = table_content_for_hash(obj)
            block_hash_key = f"table_{content_hash(table_content)}"
            block_occurrence = block_hash_counts.get(block_hash_key, 0)
            table_block_id = make_block_id("table", table_content, block_occurrence)
            block_hash_counts[block_hash_key] = block_occurrence + 1

            # Search all cells for images with hierarchical block_id
            # Use true grid coordinates to handle merged cells correctly
            rows, cols = len(obj.rows), len(obj.columns)
            for r in range(rows):
                for c in range(cols):
                    cell = obj.cell(r, c)
                    for p_idx, para in enumerate(cell.paragraphs):
                        # Hierarchical block_id: table_abc_0#r0c0/p0
                        hier_block_id = f"{table_block_id}#r{r}c{c}/p{p_idx}"
                        images.extend(
                            _extract_images_from_paragraph(
                                doc, para, hier_block_id, image_hash_counts
                            )
                        )

    return images


def _iter_all_paragraphs(doc: Document):
    """Iterate over all paragraphs in document body and tables."""
    for kind, obj, el in iter_body_blocks(doc):
        if kind == "paragraph":
            yield obj, el
        elif kind == "table":
            for row in obj.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        yield para, para._element


def _iter_all_runs_in_paragraph(para):
    """Iterate all runs in paragraph including those inside hyperlinks.

    Uses iter_inner_content() to match build_runs() indexing.
    """

    for item in para.iter_inner_content():
        if isinstance(item, Hyperlink):
            yield from item.runs
        else:  # Run
            yield item


def _find_image_in_paragraph(doc: Document, para, target_hash: str):
    """Yield each wp:inline element in paragraph matching target_hash.

    Uses iter_inner_content() traversal to match build_images() indexing.
    """
    for run in _iter_all_runs_in_paragraph(para):
        for drawing in run._element.findall(".//w:drawing", namespaces=oxml_nsmap):
            inline = drawing.find(".//wp:inline", namespaces=oxml_nsmap)
            if inline is None:
                continue
            try:
                blip = inline.graphic.graphicData.pic.blipFill.blip
            except AttributeError:
                continue
            h = get_embedded_image_hash(doc, blip)
            if h == target_hash:
                yield inline


def resolve_image(doc: Document, image_id: str) -> tuple:
    """Find embedded image by content-addressable ID. Returns (inline_el, para_el)."""
    target_hash, occurrence_str = _IMAGE_ID_RE.match(image_id).groups()
    target_occurrence = int(occurrence_str)

    occurrence_count = 0
    for para, para_el in _iter_all_paragraphs(doc):
        for inline in _find_image_in_paragraph(doc, para, target_hash):
            if occurrence_count == target_occurrence:
                return inline, para_el
            occurrence_count += 1

    raise ValueError(f"Image not found: {image_id}")


def insert_image(
    doc: Document,
    image_path: str,
    target_id: str,
    position: str,
    width_inches: float = 0,
    height_inches: float = 0,
) -> str:
    """Insert image at target location.

    Supports hierarchical target IDs:
    - table_abc_0#r0c1/p0 -> Insert into paragraph in cell
    - table_abc_0#r0c1 -> Insert into first paragraph of cell
    - table_abc_0 -> Insert before/after table
    - paragraph_abc_0 -> Insert before/after paragraph
    """

    target = resolve_target(doc, target_id)

    # Get image hash first
    _, image = doc.part.get_or_add_image(image_path)
    h = image.sha1[:8]

    width = Inches(width_inches) if width_inches else None
    height = Inches(height_inches) if height_inches else None

    if target.leaf_kind == "paragraph":
        # Insert into this paragraph
        para = target.leaf_obj
        run = para.add_run()
        run.add_picture(image_path, width, height)
        occurrence = count_image_occurrence(doc, h, para._element)
        return f"image_{h}_{occurrence}"

    if target.leaf_kind == "cell":
        # Use first paragraph of cell (create if needed)
        cell = target.leaf_obj
        para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        run = para.add_run()
        run.add_picture(image_path, width, height)
        occurrence = count_image_occurrence(doc, h, para._element)
        return f"image_{h}_{occurrence}"

    # Target is table or paragraph at base level -> insert before/after
    new_para = doc.add_paragraph()
    run = new_para.add_run()
    run.add_picture(image_path, width, height)

    if position == "before":
        target.leaf_el.addprevious(new_para._element)
    else:
        target.leaf_el.addnext(new_para._element)

    occurrence = count_image_occurrence(doc, h, new_para._element)
    return f"image_{h}_{occurrence}"


def count_image_occurrence(doc: Document, target_hash: str, target_para_el) -> int:
    """Count embedded images with same hash before target paragraph."""
    occurrence = 0
    for para, para_el in _iter_all_paragraphs(doc):
        for _inline in _find_image_in_paragraph(doc, para, target_hash):
            if para_el is target_para_el:
                return occurrence
            occurrence += 1
    raise ValueError("Target paragraph not found")


def delete_image(doc: Document, image_id: str) -> None:
    """Delete an image. Removes containing paragraph if only whitespace remains."""
    inline_el, para_el = resolve_image(doc, image_id)

    # Remove the drawing element (parent of inline)
    drawing_el = inline_el.getparent()
    drawing_el.getparent().remove(drawing_el)

    # Check if paragraph has any remaining content (text, drawings, fields, etc.)
    para = Paragraph(para_el, doc)
    has_content = bool(para.text.strip())
    if not has_content:
        for run in _iter_all_runs_in_paragraph(para):
            # Check for drawings, fields, or other significant content
            run_el = run._element
            if run_el.findall(".//w:drawing", namespaces=oxml_nsmap) or run_el.findall(
                ".//w:fldChar", namespaces=oxml_nsmap
            ):
                has_content = True
                break

    if not has_content:
        para_el.getparent().remove(para_el)
