"""Core utilities for Word document operations.

Contains shared helpers used across all ops modules:
- Content-addressed ID generation and resolution
- Block iteration and traversal
- Target resolution for hierarchical addressing
- Common XML element operations

Supports both python-docx Document and pure OOXML WordPackage.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from mcp_handley_lab.word.models import Block
from mcp_handley_lab.word.opc.constants import NSMAP, qn

if TYPE_CHECKING:
    pass

# Element tag constants for pure OOXML
_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_TR = qn("w:tr")
_W_TC = qn("w:tc")
_W_R = qn("w:r")
_W_T = qn("w:t")
_W_PPR = qn("w:pPr")
_W_PSTYLE = qn("w:pStyle")
_W_VAL = qn("w:val")

# Regexes for parsing IDs and styles
_HEADING_RE = re.compile(
    r"^Heading ?([1-9])$"
)  # Matches both "Heading1" and "Heading 1"
_ID_RE = re.compile(r"^(paragraph|heading[1-9]|table)_([0-9a-f]{8})_(\d+)$")
_IMAGE_ID_RE = re.compile(r"^image_([0-9a-f]{8})_(\d+)$")

# Hierarchical path segment patterns (0-based indices)
_CELL_RE = re.compile(r"^r(\d+)c(\d+)$")
_PARA_RE = re.compile(r"^p(\d+)$")
_TABLE_RE = re.compile(r"^tbl(\d+)$")

# Unit conversions
_EMU_PER_INCH = 914400

# PathSegment: (kind, indices) where kind='cell'|'para'|'tbl'
PathSegment = tuple[str, tuple[int, ...]]


@dataclass
class ResolvedTarget:
    """Result of resolving a hierarchical target ID.

    Pure OOXML: Works with lxml elements directly.
    Also includes wrapper objects for transitional python-docx compatibility.
    """

    base_id: str
    base_kind: str  # 'table' | 'paragraph' | 'heading1' etc
    base_el: etree._Element  # w:p or w:tbl element
    base_occurrence: int
    leaf_kind: str  # 'table' | 'cell' | 'paragraph'
    leaf_el: etree._Element  # w:p, w:tbl, or w:tc element
    # Wrapper objects for transitional python-docx compatibility (may be None)
    base_obj: any = None  # Table or Paragraph wrapper
    leaf_obj: any = None  # Table, Paragraph, or _Cell wrapper


# =============================================================================
# Content Hashing and ID Generation
# =============================================================================


def content_hash(text: str) -> str:
    """8-char SHA256 of normalized text for content-addressable IDs."""
    normalized = re.sub(r"\s+", " ", text.strip().lower())
    return hashlib.sha256(normalized.encode()).hexdigest()[:8]


def make_block_id(block_type: str, text: str, occurrence: int) -> str:
    """Generate content-addressable block ID. Single source of truth."""
    return f"{block_type}_{content_hash(text)}_{occurrence}"


# =============================================================================
# Pure OOXML Element Helpers
# =============================================================================


def get_paragraph_text(p_el: etree._Element) -> str:
    """Extract text from w:p element."""
    texts = []
    for t_el in p_el.iter(_W_T):
        if t_el.text:
            texts.append(t_el.text)
    return "".join(texts)


def get_paragraph_style(p_el: etree._Element) -> str:
    """Get paragraph style name from w:p element."""
    pPr = p_el.find(_W_PPR)
    if pPr is not None:
        pStyle = pPr.find(_W_PSTYLE)
        if pStyle is not None:
            return pStyle.get(_W_VAL, "Normal")
    return "Normal"


def get_cell_text(tc_el: etree._Element) -> str:
    """Extract text from w:tc element."""
    texts = []
    for p_el in tc_el.iter(_W_P):
        texts.append(get_paragraph_text(p_el))
    return "\n".join(texts)


def table_content_for_hash(tbl_el: etree._Element) -> str:
    """Get full table content as canonical string for hashing.

    Pure OOXML: Takes w:tbl element.
    """
    rows_data = []
    for tr_el in tbl_el.findall(_W_TR):
        cells_data = []
        for tc_el in tr_el.findall(_W_TC):
            cells_data.append(get_cell_text(tc_el))
        rows_data.append(cells_data)
    return json.dumps(rows_data, ensure_ascii=False, separators=(",", ":"))


def paragraph_kind_and_level(p_el: etree._Element) -> tuple[str, int]:
    """Detect if paragraph is a heading and return (kind, level).

    Pure OOXML: Takes w:p element.
    For content-hash IDs, kind includes level: 'heading1', 'heading2', etc.
    """
    style_name = get_paragraph_style(p_el)
    m = _HEADING_RE.match(style_name)
    if m:
        level = int(m.group(1))
        return (f"heading{level}", level)
    return ("paragraph", 0)


# =============================================================================
# Duck-Typed Helpers
# =============================================================================


def mark_dirty(pkg, partname: str = "/word/document.xml") -> None:
    """Mark an XML part as modified (only needed for WordPackage).

    Duck-typed: Safe to call with either WordPackage or Document.
    For Document, modifications to lxml elements persist automatically.
    For WordPackage, must call this after modifying cached XML so save() serializes it.
    """
    if hasattr(pkg, "mark_xml_dirty"):
        pkg.mark_xml_dirty(partname)


# =============================================================================
# Block Iteration
# =============================================================================


def iter_body_blocks(
    pkg,  # WordPackage or Document (duck-typed for transitional period)
) -> Iterator[tuple[str, etree._Element]]:
    """Yield (block_kind, element) in true document order.

    Pure OOXML: Works with WordPackage and lxml elements.
    Also accepts python-docx Document during transitional period.
    """
    # Duck-type: WordPackage has .body, Document has .element.body
    if hasattr(pkg, "body"):
        body = pkg.body
    elif hasattr(pkg, "element"):
        body = pkg.element.body
    else:
        raise TypeError(f"Expected WordPackage or Document, got {type(pkg)}")

    for child in body.iterchildren():
        if child.tag == _W_P:
            yield ("paragraph", child)
        elif child.tag == _W_TBL:
            yield ("table", child)


def _iter_all_paragraphs(pkg) -> Iterator[etree._Element]:
    """Iterate over all paragraphs in document body and tables.

    Duck-typed: Yields w:p elements.
    """
    for kind, el in iter_body_blocks(pkg):
        if kind == "paragraph":
            yield el
        elif kind == "table":
            yield from el.iter(_W_P)


def _iter_all_runs_in_paragraph(p_el: etree._Element) -> Iterator[etree._Element]:
    """Iterate all w:r elements in paragraph including those inside hyperlinks.

    Pure OOXML: Yields w:r elements.
    """
    yield from p_el.iter(_W_R)


# =============================================================================
# Target Resolution
# =============================================================================


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


def get_table_cell(tbl_el: etree._Element, row: int, col: int) -> etree._Element:
    """Get cell element at (row, col) from table element.

    Simple positional lookup. Does not account for grid spans or merges.
    """
    rows = tbl_el.findall(_W_TR)
    tr_el = rows[row]
    cells = tr_el.findall(_W_TC)
    return cells[col]


def get_cell_paragraphs(tc_el: etree._Element) -> list[etree._Element]:
    """Get all paragraph elements within a cell."""
    return tc_el.findall(_W_P)


def get_cell_tables(tc_el: etree._Element) -> list[etree._Element]:
    """Get all nested table elements within a cell."""
    return tc_el.findall(_W_TBL)


def resolve_path(
    container_el: etree._Element, segments: list[PathSegment]
) -> etree._Element:
    """Resolve path segments from container element. Returns leaf element.

    Pure OOXML: Works with lxml elements directly.
    Transition rules:
    - From Table -> cell (r,c) only
    - From Cell -> para (p) or tbl (tbl)
    - From nested Table -> cell (r,c)
    """
    current = container_el

    for kind, indices in segments:
        if kind == "cell":
            row, col = indices
            current = get_table_cell(current, row, col)
        elif kind == "para":
            current = get_cell_paragraphs(current)[indices[0]]
        elif kind == "tbl":
            current = get_cell_tables(current)[indices[0]]

    return current


def _resolve_base_block(pkg, base_id: str) -> tuple[str, etree._Element, int]:
    """Resolve base block ID to (block_type, element, occurrence).

    Duck-typed: Returns lxml element.
    Uses content-hash IDs: {type}_{hash}_{occurrence}
    Searches all blocks for matching type+hash, then skips to Nth occurrence.
    """
    match = _ID_RE.match(base_id)
    if not match:
        raise ValueError(f"Invalid block ID format: {base_id}")
    target_type, target_hash, occurrence_str = match.groups()
    target_occurrence = int(occurrence_str)

    occurrence_count = 0
    for kind, el in iter_body_blocks(pkg):
        if kind == "table":
            if target_type == "table":
                table_content = table_content_for_hash(el)
                if content_hash(table_content) == target_hash:
                    if occurrence_count == target_occurrence:
                        return "table", el, occurrence_count
                    occurrence_count += 1
        else:
            block_type, _ = paragraph_kind_and_level(el)
            if target_type == block_type:
                text = get_paragraph_text(el)
                if content_hash(text) == target_hash:
                    if occurrence_count == target_occurrence:
                        return block_type, el, occurrence_count
                    occurrence_count += 1
    raise ValueError(f"Block not found: {base_id}")


def _find_wrapper_for_element(doc, el: etree._Element):
    """Find python-docx wrapper object for an lxml element.

    Used during transitional period to populate base_obj/leaf_obj in ResolvedTarget.
    Returns Paragraph, Table, or _Cell wrapper, or None if not found.
    """

    if el.tag == _W_P:
        for para in doc.paragraphs:
            if para._element is el:
                return para
        # Also check tables for paragraphs in cells
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para._element is el:
                            return para
    elif el.tag == _W_TBL:
        for table in doc.tables:
            if table._tbl is el:
                return table
    elif el.tag == _W_TC:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell._tc is el:
                        return cell
    return None


def resolve_target(pkg, target_id: str) -> ResolvedTarget:
    """Resolve target_id to ResolvedTarget with base and leaf info.

    Duck-typed: Works with WordPackage or Document.
    Supports hierarchical IDs: table_abc12345_0#r0c1/p0
    When Document is passed, also populates base_obj/leaf_obj wrapper objects.
    """
    base_id, path_segments = parse_target_id(target_id)

    # Resolve base block
    base_kind, base_el, base_occurrence = _resolve_base_block(pkg, base_id)

    # Find wrapper object if Document is passed (for edit operations)
    base_obj = None
    is_document = hasattr(
        pkg, "paragraphs"
    )  # Document has .paragraphs, WordPackage doesn't
    if is_document:
        base_obj = _find_wrapper_for_element(pkg, base_el)

    if not path_segments:
        return ResolvedTarget(
            base_id=base_id,
            base_kind=base_kind,
            base_el=base_el,
            base_occurrence=base_occurrence,
            leaf_kind=base_kind,
            leaf_el=base_el,
            base_obj=base_obj,
            leaf_obj=base_obj,  # Same as base when no path
        )

    # Resolve path within container
    leaf_el = resolve_path(base_el, path_segments)

    # Determine leaf kind based on element tag
    if leaf_el.tag == _W_TBL:
        leaf_kind = "table"
    elif leaf_el.tag == _W_TC:
        leaf_kind = "cell"
    else:
        leaf_kind = "paragraph"

    # Find leaf wrapper object if Document is passed
    leaf_obj = None
    if is_document:
        leaf_obj = _find_wrapper_for_element(pkg, leaf_el)

    return ResolvedTarget(
        base_id=base_id,
        base_kind=base_kind,
        base_el=base_el,
        base_occurrence=base_occurrence,
        leaf_kind=leaf_kind,
        leaf_el=leaf_el,
        base_obj=base_obj,
        leaf_obj=leaf_obj,
    )


def find_paragraph_by_id(pkg, target_id: str) -> etree._Element | None:
    """Find a paragraph by its block ID.

    Duck-typed: Returns w:p element if found, None otherwise.
    """
    try:
        target = resolve_target(pkg, target_id)
        if target.leaf_kind == "paragraph":
            return target.leaf_el
        return None
    except ValueError:
        return None


def count_occurrence(
    pkg,
    block_type: str,
    text: str,
    target_el: etree._Element,  # pkg: WordPackage or Document
) -> int:
    """Count how many blocks with same type+hash appear before target_el.

    Takes WordPackage or Document (duck-typed) and lxml element.
    Used after edits to compute the correct occurrence index for returned ID.
    For tables, 'text' should be the full table content (from table_content_for_hash).
    """
    target_hash = content_hash(text)
    occurrence = 0
    for kind, el in iter_body_blocks(pkg):
        if kind == "table":
            if block_type == "table":
                table_content = table_content_for_hash(el)
                if content_hash(table_content) == target_hash:
                    if el is target_el:
                        return occurrence
                    occurrence += 1
        else:
            actual_type, _ = paragraph_kind_and_level(el)
            if block_type == actual_type:
                block_text = get_paragraph_text(el)
                if content_hash(block_text) == target_hash:
                    if el is target_el:
                        return occurrence
                    occurrence += 1
    raise ValueError("target_el not found in document")


# =============================================================================
# Block Building
# =============================================================================


def build_blocks(
    pkg,
    offset: int = 0,
    limit: int = 50,
    heading_only: bool = False,
    search_query: str = "",
) -> tuple[list[Block], int]:
    """Build list of Block objects from document body.

    Duck-typed: Works with WordPackage or Document.
    Uses content-hash IDs: {type}_{hash}_{occurrence}
    """
    # Import here to avoid circular dependency (table_to_markdown is in tables.py)
    from mcp_handley_lab.word.ops.tables import table_to_markdown

    blocks = []
    matched = 0
    query_lower = search_query.lower() if search_query else ""
    hash_counts: dict[str, int] = {}  # {type_hash: count} for occurrence tracking

    for kind, el in iter_body_blocks(pkg):
        if kind == "paragraph":
            block_type, level = paragraph_kind_and_level(el)
            text = get_paragraph_text(el)
            style = get_paragraph_style(el)

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
            md, rows, cols = table_to_markdown(el)
            # Use full content for hash (not truncated markdown)
            table_content = table_content_for_hash(el)

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
                        style="",  # Table style not easily available in pure OOXML
                        rows=rows,
                        cols=cols,
                    )
                )
            matched += 1

    return blocks, matched


# =============================================================================
# Element Operations
# =============================================================================


def make_element(
    tag: str, nsmap_override: dict | None = None, **attrs: str
) -> etree._Element:
    """Create element with namespace handling.

    Args:
        tag: Namespace-prefixed tag like "w:p"
        nsmap_override: Optional namespace map (only needed at root elements)
        **attrs: Attributes as namespace:name=value pairs
    """
    el = etree.Element(qn(tag), nsmap=nsmap_override)
    for attr, value in attrs.items():
        if ":" in attr:
            el.set(qn(attr), value)
        else:
            el.set(attr, value)
    return el


def _make_run_with(*elements) -> etree._Element:
    """Create a w:r element containing the given child elements."""
    r = etree.Element(_W_R)
    for elem in elements:
        r.append(elem)
    return r


def _insert_at(
    target_el: etree._Element, new_el: etree._Element, position: str
) -> None:
    """Insert new_el before or after target_el."""
    (target_el.addprevious if position == "before" else target_el.addnext)(new_el)


# =============================================================================
# Re-exports for convenience
# =============================================================================

# Make NSMAP available for other ops modules (replaces docx.oxml.ns.nsmap)
nsmap = NSMAP
