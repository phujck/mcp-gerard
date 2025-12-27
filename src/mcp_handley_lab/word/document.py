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

from mcp_handley_lab.word.models import Block, DocumentMeta

_HEADING_RE = re.compile(r"^Heading ([1-9])$")
_ID_RE = re.compile(r"^(paragraph|heading|table)_(\d+)_([0-9a-f]{8})$")
_WS_RE = re.compile(r"\s+")


def compute_version(file_path: str) -> str:
    """Compute SHA256 hash of word/document.xml for optimistic concurrency."""
    with zipfile.ZipFile(file_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
    return hashlib.sha256(xml_bytes).hexdigest()


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


def normalize_text(text: str) -> str:
    """Normalize text for hashing (strip, collapse whitespace)."""
    return _WS_RE.sub(" ", (text or "").strip())


def short_sha256(text: str) -> str:
    """First 8 chars of SHA256 hex digest."""
    return hashlib.sha256(text.encode()).hexdigest()[:8]


def make_block_id(block_type: str, index: int, text: str) -> str:
    """Generate stable-ish block ID: {type}_{index}_{hash[:8]}."""
    return f"{block_type}_{index}_{short_sha256(normalize_text(text))}"


def paragraph_kind_and_level(p: Paragraph) -> tuple[str, int]:
    """Detect if paragraph is a heading and return (kind, level)."""
    style_name = p.style.name if p.style else "Normal"
    m = _HEADING_RE.match(style_name)
    if m:
        return ("heading", int(m.group(1)))
    return ("paragraph", 0)


def escape_md(text: str) -> str:
    """Escape text for markdown table cells."""
    return (text or "").replace("|", "\\|").replace("\n", "<br>")


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
    """Build list of Block objects from document body."""
    blocks = []
    total = 0
    query_lower = search_query.lower() if search_query else ""

    for i, (kind, obj, _el) in enumerate(iter_body_blocks(doc)):
        total += 1

        if kind == "paragraph":
            block_type, level = paragraph_kind_and_level(obj)
            text = obj.text or ""
            style = obj.style.name if obj.style else "Normal"

            if heading_only and block_type != "heading":
                continue
            if query_lower and query_lower not in text.lower():
                continue

            if i >= offset and len(blocks) < limit:
                blocks.append(
                    Block(
                        id=make_block_id(block_type, i, text),
                        type=block_type,
                        text=text,
                        style=style,
                        level=level,
                    )
                )

        elif kind == "table":
            md, rows, cols = table_to_markdown(obj)
            style = obj.style.name if obj.style else ""

            if heading_only:
                continue
            if query_lower and query_lower not in md.lower():
                continue

            if i >= offset and len(blocks) < limit:
                blocks.append(
                    Block(
                        id=make_block_id("table", i, md),
                        type="table",
                        text=md,
                        style=style,
                        rows=rows,
                        cols=cols,
                    )
                )

    return blocks, total


def detect_toc(file_path: str, doc: Document) -> bool:
    """Detect TOC fields or styles."""
    with zipfile.ZipFile(file_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    if (
        "TOC \\" in xml
        or ("w:fldSimple" in xml and "TOC" in xml)
        or ("w:instrText" in xml and "TOC" in xml)
    ):
        return True
    for p in doc.paragraphs:
        if p.style and p.style.name.upper().startswith("TOC"):
            return True
    return False


def detect_protection(file_path: str) -> bool:
    """Detect document protection."""
    with zipfile.ZipFile(file_path, "r") as zf:
        if "word/settings.xml" not in zf.namelist():
            return False
        settings = zf.read("word/settings.xml").decode("utf-8", errors="ignore")
    return "w:documentProtection" in settings


def detect_fields(file_path: str) -> bool:
    """Detect Word fields (may affect editing)."""
    with zipfile.ZipFile(file_path, "r") as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    return "w:fldSimple" in xml or "w:instrText" in xml or "w:fldChar" in xml


def collect_warnings(file_path: str, doc: Document) -> list[str]:
    """Collect warnings about document features that may affect editing."""
    warnings = []
    if detect_protection(file_path):
        warnings.append("Document is protected. Edits may fail.")
    if detect_toc(file_path, doc):
        warnings.append(
            "Document contains TOC. TOC may need manual refresh in Word after edits."
        )
    if detect_fields(file_path):
        warnings.append(
            "Document contains Word fields. Editing nearby content may affect field integrity."
        )
    return warnings


def resolve_target(
    doc: Document, target_id: str
) -> tuple[str, Paragraph | Table, CT_P | CT_Tbl, int]:
    """Resolve target_id to (kind, obj, element, body_index)."""
    m = _ID_RE.match(target_id)
    if not m:
        raise ValueError(f"Invalid target_id format: {target_id}")
    declared_type, idx_s, hash8 = m.groups()
    idx = int(idx_s)

    for i, (kind, obj, el) in enumerate(iter_body_blocks(doc)):
        if i == idx:
            # Validate declared type matches actual type
            if kind == "table":
                if declared_type != "table":
                    raise ValueError(
                        f"Type mismatch: ID declares '{declared_type}' but block at index {idx} is a table."
                    )
                md, _, _ = table_to_markdown(obj)
                expected_id = make_block_id("table", i, md)
            else:
                actual_type, _ = paragraph_kind_and_level(obj)
                if declared_type != actual_type:
                    raise ValueError(
                        f"Type mismatch: ID declares '{declared_type}' but block at index {idx} is a {actual_type}."
                    )
                expected_id = make_block_id(actual_type, i, obj.text or "")
            if not expected_id.endswith(hash8):
                raise ValueError(
                    f"Block hash mismatch at index {idx}. Document may have changed. Re-read required."
                )
            return kind, obj, el, i

    raise ValueError(f"Block index out of range: {idx}")


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


def delete_block(kind: str, obj: Paragraph | Table) -> None:
    """Delete a block from the document."""
    el = obj._element if kind == "paragraph" else obj._tbl
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
    """Create a new block (paragraph, heading, table, page_break) and return (kind, obj, element)."""
    if content_type == "paragraph":
        p = doc.add_paragraph(content_data)
        if style_name:
            p.style = style_name
        return ("paragraph", p, p._element)
    elif content_type == "heading":
        level = max(1, min(heading_level, 9))
        p = doc.add_heading(content_data, level=level)
        return ("paragraph", p, p._element)
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
    elif content_type == "page_break":
        doc.add_page_break()
        last_p = doc.paragraphs[-1]
        return ("paragraph", last_p, last_p._element)
    elif content_type == "list":
        list_data = json.loads(content_data)
        items = list_data.get("items", [])
        p = doc.add_paragraph()
        p.text = "\n".join(f"• {item}" for item in items)
        return ("paragraph", p, p._element)
    elif content_type == "image":
        p = doc.add_paragraph()
        run = p.add_run()
        run.add_picture(content_data)
        return ("paragraph", p, p._element)
    else:
        raise ValueError(f"Unknown content_type: {content_type}")
