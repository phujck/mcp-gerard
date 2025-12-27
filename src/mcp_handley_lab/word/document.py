"""Python-docx wrapper functions for Word document manipulation."""

import contextlib
import hashlib
import json
import re
import zipfile
from collections.abc import Iterator

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from mcp_handley_lab.word.models import (
    Block,
    CellInfo,
    CommentInfo,
    DocumentMeta,
    RunInfo,
)

_HEADING_RE = re.compile(r"^Heading ([1-9])$")
_ID_RE = re.compile(r"^(paragraph|heading[1-9]|table)_([0-9a-f]{8})_(\d+)$")


def content_hash(text: str) -> str:
    """8-char SHA256 of normalized text for content-addressable IDs."""
    normalized = re.sub(r"\s+", " ", (text or "").strip().lower())
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


def make_block_id(block_type: str, text: str, occurrence: int) -> str:
    """Generate content-addressable block ID: {type}_{hash}_{occurrence}."""
    return f"{block_type}_{content_hash(text)}_{occurrence}"


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


def escape_md(text: str) -> str:
    """Escape text for markdown table cells."""
    return (text or "").replace("|", "\\|").replace("\n", "<br>")


def table_content_for_hash(table: Table) -> str:
    """Get full table content as canonical string for hashing.

    Returns JSON-like representation of all cells for content-addressing.
    This ensures large tables get unique hashes even if preview is truncated.
    """
    rows = []
    for row in table.rows:
        cells = [cell.text for cell in row.cells]
        rows.append(cells)
    return json.dumps(rows, ensure_ascii=False, separators=(",", ":"))


def table_to_markdown(
    table: Table, max_chars: int = 500, max_rows: int = 20, max_cols: int = 10
) -> tuple[str, int, int]:
    """Convert table to markdown preview with truncation."""
    rows = len(table.rows)
    cols = len(table.columns) if rows else 0
    if not rows or not cols:
        return "| |\\n|---|", rows, cols

    r_lim = min(rows, max_rows)
    c_lim = min(cols, max_cols)

    grid = []
    for r in range(r_lim):
        row = [escape_md(table.cell(r, c).text.strip()) for c in range(c_lim)]
        grid.append(row)

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
            if query_lower and query_lower not in md.lower():
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


def collect_warnings(file_path: str) -> list[str]:
    """Collect warnings about document features that may affect editing."""
    # Simple check for complex features that may not round-trip
    with zipfile.ZipFile(file_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    has_complex = (
        "w:fldSimple" in xml
        or "w:instrText" in xml
        or "w:documentProtection" in xml
        or "TOC" in xml
    )
    if has_complex:
        return ["Complex Word features detected; verify in Word after edits."]
    return []


def resolve_target(
    doc: Document, target_id: str
) -> tuple[str, Paragraph | Table, CT_P | CT_Tbl, int]:
    """Resolve target_id to (block_type, obj, element, occurrence).

    Uses content-hash IDs: {type}_{hash}_{occurrence}
    Searches all blocks for matching type+hash, then skips to Nth occurrence.
    """
    m = _ID_RE.match(target_id)
    if not m:
        raise ValueError(f"Invalid target_id format: {target_id}")
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
    return occurrence  # fallback


def insert_paragraph_relative(
    doc: Document,
    target_el: CT_P | CT_Tbl,
    text: str,
    position: str,
    style_name: str = "",
) -> Paragraph:
    """Insert paragraph before/after target element."""
    new_p = doc.add_paragraph(text)
    if style_name:
        new_p.style = style_name
    if position == "before":
        target_el.addprevious(new_p._element)
    else:
        target_el.addnext(new_p._element)
    return new_p


def insert_heading_relative(
    doc: Document, target_el: CT_P | CT_Tbl, text: str, level: int, position: str
) -> Paragraph:
    """Insert heading before/after target element."""
    level = max(1, min(level, 9))
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
    rows = len(table_data)
    cols = max((len(r) for r in table_data), default=1)
    tbl = doc.add_table(rows=rows, cols=cols)
    if style_name:
        with contextlib.suppress(KeyError):
            tbl.style = style_name
    for r in range(rows):
        for c in range(len(table_data[r])):
            if c < cols:
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


def replace_paragraph_text(p: Paragraph, new_text: str) -> None:
    """Replace paragraph text (destroys run-level formatting)."""
    p.text = new_text


def replace_table(doc: Document, old_tbl: Table, table_data: list[list[str]]) -> Table:
    """Replace table with new data."""
    old_el = old_tbl._tbl
    rows = len(table_data)
    cols = max((len(r) for r in table_data), default=1)
    new_tbl = doc.add_table(rows=rows, cols=cols)
    for r in range(rows):
        for c in range(len(table_data[r])):
            if c < cols:
                new_tbl.cell(r, c).text = str(table_data[r][c])
    old_el.addprevious(new_tbl._tbl)
    old_el.getparent().remove(old_el)
    return new_tbl


def apply_paragraph_style(p: Paragraph, style_name: str) -> None:
    """Apply Word style to paragraph."""
    p.style = style_name


def apply_paragraph_formatting(p: Paragraph, fmt: dict) -> None:
    """Apply direct formatting to paragraph (affects all runs)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    align = fmt.get("alignment")
    if align:
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if align.lower() in alignment_map:
            p.alignment = alignment_map[align.lower()]

    if not p.runs:
        p.add_run(p.text or "")

    for run in p.runs:
        if "bold" in fmt:
            run.bold = bool(fmt["bold"])
        if "italic" in fmt:
            run.italic = bool(fmt["italic"])
        if "underline" in fmt:
            run.underline = bool(fmt["underline"])
        if "font_name" in fmt:
            run.font.name = fmt["font_name"]
        if "font_size" in fmt:
            run.font.size = Pt(float(fmt["font_size"]))
        if "color" in fmt:
            color = fmt["color"].lstrip("#")
            run.font.color.rgb = RGBColor.from_string(color)


def find_body_index_of_element(doc: Document, element: CT_P | CT_Tbl) -> int | None:
    """Find body index of an element by identity."""
    for i, child in enumerate(doc.element.body.iterchildren()):
        if child is element:
            return i
    return None


def create_block(
    doc: Document,
    content_type: str,
    content_data: str,
    heading_level: int = 1,
    style_name: str = "",
) -> tuple[str, Paragraph | Table, CT_P | CT_Tbl]:
    """Create a new block (paragraph, heading, table) and return (block_type, obj, element)."""
    if content_type == "paragraph":
        p = doc.add_paragraph(content_data)
        if style_name:
            p.style = style_name
        return ("paragraph", p, p._element)
    elif content_type == "heading":
        level = max(1, min(heading_level, 9))
        p = doc.add_heading(content_data, level=level)
        return (f"heading{level}", p, p._element)
    elif content_type == "table":
        table_data = json.loads(content_data)
        rows = len(table_data)
        cols = max((len(r) for r in table_data), default=1)
        tbl = doc.add_table(rows=rows, cols=cols)
        if style_name:
            with contextlib.suppress(KeyError):
                tbl.style = style_name
        for r in range(rows):
            for c in range(len(table_data[r])):
                if c < cols:
                    tbl.cell(r, c).text = str(table_data[r][c])
        return ("table", tbl, tbl._tbl)
    else:
        raise ValueError(
            f"Unknown content_type: {content_type}. Use 'paragraph', 'heading', or 'table'."
        )


def build_table_cells(table: Table) -> list[CellInfo]:
    """Build list of CellInfo for all cells in a table."""
    cells = []
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            cells.append(
                CellInfo(
                    row=r_idx + 1,  # 1-based
                    col=c_idx + 1,  # 1-based
                    text=cell.text or "",
                )
            )
    return cells


def replace_table_cell(table: Table, row: int, col: int, text: str) -> None:
    """Replace text in a table cell. Row/col are 1-based."""
    r_idx = row - 1
    c_idx = col - 1
    if r_idx < 0 or r_idx >= len(table.rows):
        raise ValueError(f"Row {row} out of range (table has {len(table.rows)} rows)")
    if c_idx < 0 or c_idx >= len(table.columns):
        raise ValueError(
            f"Column {col} out of range (table has {len(table.columns)} columns)"
        )
    table.cell(r_idx, c_idx).text = text


def has_hyperlinks(paragraph: Paragraph) -> bool:
    """Check if paragraph contains hyperlinks (runs inside hyperlinks not in Paragraph.runs)."""
    return len(paragraph._p.hyperlink_lst) > 0


def has_complex_content(run) -> bool:
    """Check if run contains non-text elements that would be deleted on text edit."""
    r = run._r
    for child in r:
        tag = child.tag
        if (
            tag.endswith("}drawing")
            or tag.endswith("}lastRenderedPageBreak")
            or tag.endswith("}object")
        ):
            return True
    return False


def build_runs(paragraph: Paragraph) -> tuple[list[RunInfo], list[str]]:
    """Build list of RunInfo for all runs in a paragraph.

    Returns (runs, warnings). Only includes direct runs, not runs inside hyperlinks.
    """
    runs = []
    warnings = []

    if has_hyperlinks(paragraph):
        warnings.append(
            "Paragraph contains hyperlinks; runs scope shows direct runs only"
        )

    for idx, run in enumerate(paragraph.runs):
        # Extract formatting properties
        font = run.font
        color_hex = None
        if font.color and font.color.rgb:
            color_hex = str(font.color.rgb)

        font_size = None
        if font.size:
            font_size = font.size.pt

        complex_content = has_complex_content(run)
        if complex_content:
            warnings.append(
                f"Run {idx} contains non-text content; editing text will remove it"
            )

        runs.append(
            RunInfo(
                index=idx,
                text=run.text or "",
                bold=run.bold,
                italic=run.italic,
                underline=run.underline,
                font_name=font.name,
                font_size=font_size,
                color=color_hex,
                has_complex_content=complex_content,
            )
        )

    return runs, warnings


def edit_run_text(paragraph: Paragraph, run_index: int, text: str) -> list[str]:
    """Edit run text. Returns warnings if run had complex content."""
    warnings = []
    if run_index < 0 or run_index >= len(paragraph.runs):
        raise ValueError(
            f"Run index {run_index} out of range (paragraph has {len(paragraph.runs)} runs)"
        )

    run = paragraph.runs[run_index]
    if has_complex_content(run):
        warnings.append(
            f"Run {run_index} contained non-text content that was removed by text edit"
        )

    run.text = text
    return warnings


def edit_run_formatting(paragraph: Paragraph, run_index: int, fmt: dict) -> None:
    """Apply formatting to a specific run."""
    if run_index < 0 or run_index >= len(paragraph.runs):
        raise ValueError(
            f"Run index {run_index} out of range (paragraph has {len(paragraph.runs)} runs)"
        )

    run = paragraph.runs[run_index]

    if "bold" in fmt:
        run.bold = bool(fmt["bold"])
    if "italic" in fmt:
        run.italic = bool(fmt["italic"])
    if "underline" in fmt:
        run.underline = bool(fmt["underline"])
    if "font_name" in fmt:
        run.font.name = fmt["font_name"]
    if "font_size" in fmt:
        run.font.size = Pt(float(fmt["font_size"]))
    if "color" in fmt:
        color = fmt["color"].lstrip("#")
        run.font.color.rgb = RGBColor.from_string(color)


def build_comments(doc: Document) -> list[CommentInfo]:
    """Build list of CommentInfo from document comments."""
    comments = []
    for comment in doc.comments:
        comments.append(
            CommentInfo(
                id=comment.comment_id,
                author=comment.author,
                initials=comment.initials,
                timestamp=comment.timestamp.isoformat() if comment.timestamp else None,
                text=comment.text,
            )
        )
    return comments


def add_comment_to_block(
    doc: Document,
    paragraph: Paragraph,
    text: str,
    author: str = "",
    initials: str = "",
) -> int:
    """Add a comment anchored to all runs in a paragraph. Returns comment_id."""
    runs = paragraph.runs
    if not runs:
        # Create empty run as anchor (don't copy paragraph.text to avoid duplication)
        paragraph.add_run("")
        runs = paragraph.runs

    comment = doc.add_comment(runs=runs, text=text, author=author, initials=initials)
    return comment.comment_id
