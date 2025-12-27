"""Word document MCP tool - read and edit operations."""

import json

from docx import Document
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.word.document import (
    add_comment_to_block,
    apply_paragraph_formatting,
    build_blocks,
    build_comments,
    build_headers_footers,
    build_page_setup,
    build_runs,
    build_table_cells,
    content_hash,
    count_occurrence,
    delete_block,
    edit_run_formatting,
    edit_run_text,
    get_document_meta,
    insert_heading_relative,
    insert_paragraph_relative,
    insert_table_relative,
    replace_table,
    replace_table_cell,
    resolve_target,
    set_footer_text,
    set_header_text,
    set_page_margins,
    set_page_orientation,
    table_content_for_hash,
)
from mcp_handley_lab.word.models import DocumentReadResult, EditResult

mcp = FastMCP("Word Document Tool")


@mcp.tool(
    description="Read Word document content. Scopes: 'meta' (doc info), 'outline' (headings only), 'blocks' (all content), 'search' (find text), 'table_cells' (cells of a table), 'runs' (text runs in a paragraph), 'comments' (all comments), 'headers_footers' (headers/footers per section), 'page_setup' (margins, orientation per section). Block IDs are content-addressed (type_hash_occurrence) and stable across structural edits."
)
def read(
    file_path: str = Field(..., description="Path to .docx file"),
    scope: str = Field(
        "outline",
        description="What to read: 'meta', 'outline', 'blocks', 'search', 'table_cells', 'runs', 'comments', 'headers_footers', 'page_setup'",
    ),
    target_id: str = Field("", description="Block ID for table_cells/runs scopes"),
    search_query: str = Field("", description="Text to search for (scope='search')"),
    limit: int = Field(50, description="Max blocks to return"),
    offset: int = Field(0, description="Pagination offset"),
) -> DocumentReadResult:
    """Read Word document content."""
    doc = Document(file_path)

    if scope == "meta":
        meta = get_document_meta(doc)
        _, block_count = build_blocks(doc, offset=0, limit=0)
        return DocumentReadResult(block_count=block_count, meta=meta)
    if scope == "outline":
        blocks, block_count = build_blocks(
            doc, offset=offset, limit=limit, heading_only=True
        )
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "blocks":
        blocks, block_count = build_blocks(doc, offset=offset, limit=limit)
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "search":
        blocks, block_count = build_blocks(
            doc, offset=offset, limit=limit, search_query=search_query
        )
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "table_cells":
        block_type, obj, _, _ = resolve_target(doc, target_id)
        if block_type != "table":
            raise ValueError("target_id must be a table")
        cells = build_table_cells(obj)
        return DocumentReadResult(
            block_count=len(cells),
            cells=cells,
            table_rows=len(obj.rows),
            table_cols=len(obj.columns),
        )
    if scope == "runs":
        block_type, obj, _, _ = resolve_target(doc, target_id)
        if block_type == "table":
            raise ValueError("target_id must be a paragraph or heading")
        runs = build_runs(obj)
        return DocumentReadResult(block_count=len(runs), runs=runs)
    if scope == "comments":
        comments = build_comments(doc)
        return DocumentReadResult(block_count=len(comments), comments=comments)
    if scope == "headers_footers":
        headers_footers = build_headers_footers(doc)
        return DocumentReadResult(
            block_count=len(headers_footers), headers_footers=headers_footers
        )
    if scope == "page_setup":
        page_setup = build_page_setup(doc)
        return DocumentReadResult(block_count=len(page_setup), page_setup=page_setup)
    raise ValueError(f"Unknown scope: {scope}")


@mcp.tool(
    description="Edit Word document. Operations: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'add_comment', 'set_header', 'set_footer', 'set_margins', 'set_orientation'. Block IDs are content-addressed and stable across structural edits. Returns new ID after content changes."
)
def edit(
    file_path: str = Field(..., description="Path to .docx file"),
    operation: str = Field(
        ...,
        description="Operation: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'add_comment', 'set_header', 'set_footer', 'set_margins', 'set_orientation'",
    ),
    target_id: str = Field(
        "",
        description="Block ID from read() - required for insert/delete/replace/style/edit_cell/edit_run/add_comment",
    ),
    content_type: str = Field(
        "paragraph",
        description="Type: 'paragraph', 'heading', 'table'",
    ),
    content_data: str = Field(
        "",
        description="Content: text or JSON (for tables). For add_comment/set_header/set_footer: the text content. For set_orientation: 'portrait' or 'landscape'.",
    ),
    style_name: str = Field(
        "", description="Apply Word style: 'Heading 1', 'Normal', etc."
    ),
    formatting: str = Field(
        "",
        description='Direct formatting JSON: {"bold": true, "color": "FF0000", "font_size": 14}. For set_margins: {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25} in inches.',
    ),
    heading_level: int = Field(
        1, description="Heading level 1-9 (only for content_type='heading')"
    ),
    row: int = Field(0, description="Row number (1-based, for edit_cell operation)"),
    col: int = Field(0, description="Column number (1-based, for edit_cell operation)"),
    run_index: int = Field(
        -1,
        description="Run index (0-based, for edit_run operation). Use read() with scope='runs' to find indices.",
    ),
    author: str = Field("", description="Comment author name (for add_comment)"),
    initials: str = Field("", description="Comment author initials (for add_comment)"),
    section_index: int = Field(
        0,
        description="Section index (0-based, for set_header/set_footer/set_margins/set_orientation). Use read() with scope='page_setup' to see sections.",
    ),
) -> EditResult:
    """Edit Word document."""
    if operation == "create":
        doc = Document()
        if content_type == "paragraph":
            obj = doc.add_paragraph(content_data, style_name or None)
            block_type, el = "paragraph", obj._element
        elif content_type == "heading":
            obj = doc.add_heading(content_data, level=heading_level)
            block_type, el = f"heading{heading_level}", obj._element
        elif content_type == "table":
            table_data = json.loads(content_data)
            rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
            obj = doc.add_table(rows=rows, cols=cols)
            obj.style = style_name or "Table Grid"
            for r in range(rows):
                for c in range(len(table_data[r])):
                    obj.cell(r, c).text = str(table_data[r][c])
            block_type, el = "table", obj._tbl
        else:
            raise ValueError(f"Unknown content_type: {content_type}")
        text = table_content_for_hash(obj) if block_type == "table" else obj.text or ""
        occurrence = count_occurrence(doc, block_type, text, el)
        element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Created: {file_path}"
        )

    doc = Document(file_path)

    if operation in ("insert_before", "insert_after"):
        _, obj, target_el, _ = resolve_target(doc, target_id)
        position = "before" if operation == "insert_before" else "after"

        if content_type == "paragraph":
            new_p = insert_paragraph_relative(
                doc, target_el, content_data, position, style_name
            )
            text = new_p.text or ""
            occurrence = count_occurrence(doc, "paragraph", text, new_p._element)
            element_id = f"paragraph_{content_hash(text)}_{occurrence}"
        elif content_type == "heading":
            new_p = insert_heading_relative(
                doc, target_el, content_data, heading_level, position
            )
            text = new_p.text or ""
            block_type = f"heading{heading_level}"
            occurrence = count_occurrence(doc, block_type, text, new_p._element)
            element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        elif content_type == "table":
            table_data = json.loads(content_data)
            new_tbl = insert_table_relative(
                doc, target_el, table_data, position, style_name or "Table Grid"
            )
            table_content = table_content_for_hash(new_tbl)
            occurrence = count_occurrence(doc, "table", table_content, new_tbl._tbl)
            element_id = f"table_{content_hash(table_content)}_{occurrence}"
        else:
            raise ValueError(f"Unknown content_type: {content_type}")
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Inserted {content_type} {position} {target_id}",
        )

    if operation == "append":
        if content_type == "paragraph":
            obj = doc.add_paragraph(content_data, style_name or None)
            block_type, el = "paragraph", obj._element
        elif content_type == "heading":
            obj = doc.add_heading(content_data, level=heading_level)
            block_type, el = f"heading{heading_level}", obj._element
        elif content_type == "table":
            table_data = json.loads(content_data)
            rows, cols = len(table_data), max((len(r) for r in table_data), default=1)
            obj = doc.add_table(rows=rows, cols=cols)
            obj.style = style_name or "Table Grid"
            for r in range(rows):
                for c in range(len(table_data[r])):
                    obj.cell(r, c).text = str(table_data[r][c])
            block_type, el = "table", obj._tbl
        else:
            raise ValueError(f"Unknown content_type: {content_type}")
        text = table_content_for_hash(obj) if block_type == "table" else obj.text or ""
        occurrence = count_occurrence(doc, block_type, text, el)
        element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Appended {content_type} to document",
        )

    if operation == "delete":
        block_type, obj, _, _ = resolve_target(doc, target_id)
        delete_block(block_type, obj)
        doc.save(file_path)
        return EditResult(
            success=True, element_id=target_id, message=f"Deleted block {target_id}"
        )

    if operation == "replace":
        block_type, obj, target_el, _ = resolve_target(doc, target_id)
        if block_type.startswith("heading") or block_type == "paragraph":
            obj.text = content_data
            text = content_data
            occurrence = count_occurrence(doc, block_type, text, target_el)
            element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        elif block_type == "table":
            table_data = json.loads(content_data)
            new_tbl = replace_table(doc, obj, table_data)
            table_content = table_content_for_hash(new_tbl)
            occurrence = count_occurrence(doc, "table", table_content, new_tbl._tbl)
            element_id = f"table_{content_hash(table_content)}_{occurrence}"
        else:
            raise ValueError(f"Unsupported block_type: {block_type}")
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Replaced content of {target_id}",
        )

    if operation == "style":
        block_type, obj, target_el, _ = resolve_target(doc, target_id)
        if style_name:
            obj.style = style_name
        if block_type == "table":
            table_content = table_content_for_hash(obj)
            occurrence = count_occurrence(doc, "table", table_content, target_el)
            element_id = f"table_{content_hash(table_content)}_{occurrence}"
        else:
            if formatting:
                fmt = json.loads(formatting)
                apply_paragraph_formatting(obj, fmt)
            text = obj.text or ""
            occurrence = count_occurrence(doc, block_type, text, target_el)
            element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Applied style to {target_id}"
        )

    if operation == "edit_cell":
        block_type, obj, target_el, _ = resolve_target(doc, target_id)
        if block_type != "table":
            raise ValueError("target_id must be a table")
        replace_table_cell(obj, row, col, content_data)
        table_content = table_content_for_hash(obj)
        occurrence = count_occurrence(doc, "table", table_content, target_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Updated cell r{row}c{col}"
        )

    if operation == "edit_run":
        block_type, obj, target_el, _ = resolve_target(doc, target_id)
        if block_type == "table":
            raise ValueError("target_id must be a paragraph or heading")
        if content_data:
            edit_run_text(obj, run_index, content_data)
        if formatting:
            fmt = json.loads(formatting)
            edit_run_formatting(obj, run_index, fmt)
        text = obj.text or ""
        occurrence = count_occurrence(doc, block_type, text, target_el)
        element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Updated run {run_index}"
        )

    if operation == "add_comment":
        _, obj, _, _ = resolve_target(doc, target_id)
        comment_id = add_comment_to_block(doc, obj, content_data, author, initials)
        doc.save(file_path)
        return EditResult(
            success=True,
            comment_id=comment_id,
            message=f"Added comment {comment_id} to {target_id}",
        )

    if operation == "set_header":
        set_header_text(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set header for section {section_index}"
        )

    if operation == "set_footer":
        set_footer_text(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set footer for section {section_index}"
        )

    if operation == "set_margins":
        margins = json.loads(formatting)
        set_page_margins(
            doc,
            section_index,
            top=margins["top"],
            bottom=margins["bottom"],
            left=margins["left"],
            right=margins["right"],
        )
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set margins for section {section_index}"
        )

    if operation == "set_orientation":
        set_page_orientation(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True,
            message=f"Set orientation to {content_data} for section {section_index}",
        )
    raise ValueError(f"Unknown operation: {operation}")


if __name__ == "__main__":
    mcp.run()
