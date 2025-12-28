"""Word document MCP tool - read and edit operations."""

import json

from docx import Document
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.word.document import (
    add_break_after,
    add_comment_to_block,
    add_page_break,
    add_section,
    add_table_column,
    add_table_row,
    apply_paragraph_formatting,
    build_blocks,
    build_comments,
    build_headers_footers,
    build_images,
    build_page_setup,
    build_runs,
    build_table_cells,
    content_hash,
    count_occurrence,
    delete_block,
    delete_image,
    delete_table_column,
    delete_table_row,
    edit_run_formatting,
    edit_run_text,
    get_document_meta,
    insert_heading_relative,
    insert_image,
    insert_paragraph_relative,
    insert_table_relative,
    replace_table,
    replace_table_cell,
    resolve_target,
    set_document_meta,
    set_even_page_footer,
    set_even_page_header,
    set_first_page_footer,
    set_first_page_header,
    set_footer_text,
    set_header_text,
    set_page_margins,
    set_page_orientation,
    table_content_for_hash,
)
from mcp_handley_lab.word.models import DocumentReadResult, EditResult

mcp = FastMCP("Word Document Tool")


@mcp.tool(
    description="Read Word document content. Scopes: 'meta' (doc info), 'outline' (headings only), 'blocks' (all content), 'search' (find text), 'table_cells' (cells of a table), 'runs' (text runs in a paragraph), 'comments' (all comments), 'headers_footers' (headers/footers per section), 'page_setup' (margins, orientation per section), 'images' (embedded inline images). Block IDs are content-addressed (type_hash_occurrence) and stable across structural edits."
)
def read(
    file_path: str = Field(..., description="Path to .docx file"),
    scope: str = Field(
        "outline",
        description="What to read: 'meta', 'outline', 'blocks', 'search', 'table_cells', 'runs', 'comments', 'headers_footers', 'page_setup', 'images'",
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
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        cells = build_table_cells(t.leaf_obj, t.base_id)
        return DocumentReadResult(
            block_count=len(cells),
            cells=cells,
            table_rows=len(t.leaf_obj.rows),
            table_cols=len(t.leaf_obj.columns),
        )
    if scope == "runs":
        t = resolve_target(doc, target_id)
        if t.leaf_kind == "table":
            raise ValueError("target_id must be a paragraph or heading")
        runs = build_runs(t.leaf_obj)
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
    if scope == "images":
        images = build_images(doc)
        return DocumentReadResult(block_count=len(images), images=images)
    raise ValueError(f"Unknown scope: {scope}")


@mcp.tool(
    description="Edit Word document. Operations: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'add_comment', 'set_header', 'set_footer', 'set_first_page_header', 'set_first_page_footer', 'set_even_page_header', 'set_even_page_footer', 'set_margins', 'set_orientation', 'insert_image', 'delete_image', 'add_row', 'add_column', 'delete_row', 'delete_column', 'add_page_break', 'add_break', 'set_meta', 'add_section'. Block IDs are content-addressed and stable across structural edits. Returns new ID after content changes."
)
def edit(
    file_path: str = Field(..., description="Path to .docx file"),
    operation: str = Field(
        ...,
        description="Operation: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'add_comment', 'set_header', 'set_footer', 'set_first_page_header', 'set_first_page_footer', 'set_even_page_header', 'set_even_page_footer', 'set_margins', 'set_orientation', 'insert_image', 'delete_image', 'add_row', 'add_column', 'delete_row', 'delete_column', 'add_page_break', 'add_break', 'set_meta', 'add_section'",
    ),
    target_id: str = Field(
        "",
        description="Block ID from read() - required for insert/delete/replace/style/edit_cell/edit_run/add_comment/insert_image. Supports hierarchical IDs for insert_image into table cells: 'table_abc_0#r0c1' or 'table_abc_0#r0c1/p0' (0-based indices). For delete_image: the image ID.",
    ),
    content_type: str = Field(
        "paragraph",
        description="Type: 'paragraph', 'heading', 'table'",
    ),
    content_data: str = Field(
        "",
        description="Content: text or JSON (for tables). For add_comment/set_header/set_footer: the text content. For set_orientation: 'portrait' or 'landscape'. For insert_image: the image file path.",
    ),
    style_name: str = Field(
        "", description="Apply Word style: 'Heading 1', 'Normal', etc."
    ),
    formatting: str = Field(
        "",
        description='Direct formatting JSON. Text: {"bold": true, "color": "FF0000", "font_size": 14, "highlight_color": "yellow", "strike": true, "subscript": true, "superscript": true}. Paragraph: {"left_indent": 0.5, "right_indent": 0.5, "first_line_indent": 0.5, "space_before": 12, "space_after": 12, "line_spacing": 1.5, "keep_with_next": true, "page_break_before": true} (indents in inches, spacing in points, line_spacing < 5 is multiplier). Margins: {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25} in inches. Images: {"width": 4} or {"height": 3} in inches.',
    ),
    heading_level: int = Field(
        1, description="Heading level 1-9 (only for content_type='heading')"
    ),
    row: int = Field(0, description="Row number (0-based, for edit_cell operation)"),
    col: int = Field(0, description="Column number (0-based, for edit_cell operation)"),
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
        t = resolve_target(doc, target_id)
        position = "before" if operation == "insert_before" else "after"

        if content_type == "paragraph":
            new_p = insert_paragraph_relative(
                doc, t.leaf_el, content_data, position, style_name
            )
            text = new_p.text or ""
            occurrence = count_occurrence(doc, "paragraph", text, new_p._element)
            element_id = f"paragraph_{content_hash(text)}_{occurrence}"
        elif content_type == "heading":
            new_p = insert_heading_relative(
                doc, t.leaf_el, content_data, heading_level, position
            )
            text = new_p.text or ""
            block_type = f"heading{heading_level}"
            occurrence = count_occurrence(doc, block_type, text, new_p._element)
            element_id = f"{block_type}_{content_hash(text)}_{occurrence}"
        elif content_type == "table":
            table_data = json.loads(content_data)
            new_tbl = insert_table_relative(
                doc, t.leaf_el, table_data, position, style_name or "Table Grid"
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
        t = resolve_target(doc, target_id)
        delete_block(t.leaf_kind, t.leaf_obj)
        doc.save(file_path)
        return EditResult(
            success=True, element_id=target_id, message=f"Deleted block {target_id}"
        )

    if operation == "replace":
        t = resolve_target(doc, target_id)
        if t.leaf_kind.startswith("heading") or t.leaf_kind == "paragraph":
            t.leaf_obj.text = content_data
            text = content_data
            occurrence = count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
            element_id = f"{t.leaf_kind}_{content_hash(text)}_{occurrence}"
        elif t.leaf_kind == "table":
            table_data = json.loads(content_data)
            new_tbl = replace_table(doc, t.leaf_obj, table_data)
            table_content = table_content_for_hash(new_tbl)
            occurrence = count_occurrence(doc, "table", table_content, new_tbl._tbl)
            element_id = f"table_{content_hash(table_content)}_{occurrence}"
        else:
            raise ValueError(f"Unsupported leaf_kind: {t.leaf_kind}")
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Replaced content of {target_id}",
        )

    if operation == "style":
        t = resolve_target(doc, target_id)
        if style_name:
            t.leaf_obj.style = style_name
        if t.leaf_kind == "table":
            table_content = table_content_for_hash(t.leaf_obj)
            occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
            element_id = f"table_{content_hash(table_content)}_{occurrence}"
        else:
            if formatting:
                fmt = json.loads(formatting)
                apply_paragraph_formatting(t.leaf_obj, fmt)
            text = t.leaf_obj.text or ""
            occurrence = count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
            element_id = f"{t.leaf_kind}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Applied style to {target_id}"
        )

    if operation == "edit_cell":
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        replace_table_cell(t.leaf_obj, row, col, content_data)
        table_content = table_content_for_hash(t.leaf_obj)
        occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Updated cell r{row}c{col}"
        )

    if operation == "edit_run":
        t = resolve_target(doc, target_id)
        if t.leaf_kind == "table":
            raise ValueError("target_id must be a paragraph or heading")
        if content_data:
            edit_run_text(t.leaf_obj, run_index, content_data)
        if formatting:
            fmt = json.loads(formatting)
            edit_run_formatting(t.leaf_obj, run_index, fmt)
        text = t.leaf_obj.text or ""
        occurrence = count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
        element_id = f"{t.leaf_kind}_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Updated run {run_index}"
        )

    if operation == "add_comment":
        t = resolve_target(doc, target_id)
        comment_id = add_comment_to_block(
            doc, t.leaf_obj, content_data, author, initials
        )
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

    if operation == "set_first_page_header":
        set_first_page_header(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set first page header for section {section_index}"
        )

    if operation == "set_first_page_footer":
        set_first_page_footer(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set first page footer for section {section_index}"
        )

    if operation == "set_even_page_header":
        set_even_page_header(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set even page header for section {section_index}"
        )

    if operation == "set_even_page_footer":
        set_even_page_footer(doc, section_index, content_data)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Set even page footer for section {section_index}"
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

    if operation == "insert_image":
        fmt = json.loads(formatting) if formatting else {}
        element_id = insert_image(
            doc,
            content_data,
            target_id,
            "after",
            width_inches=float(fmt.get("width", 0)),
            height_inches=float(fmt.get("height", 0)),
        )
        doc.save(file_path)
        return EditResult(success=True, element_id=element_id, message="Inserted image")

    if operation == "delete_image":
        delete_image(doc, target_id)
        doc.save(file_path)
        return EditResult(success=True, message=f"Deleted {target_id}")

    if operation == "add_row":
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        data = json.loads(content_data) if content_data else None
        row_idx = add_table_row(t.leaf_obj, data)
        table_content = table_content_for_hash(t.leaf_obj)
        occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Added row {row_idx}"
        )

    if operation == "add_column":
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        fmt = json.loads(formatting) if formatting else {}
        width = float(fmt.get("width", 1.0))
        data = json.loads(content_data) if content_data else None
        col_idx = add_table_column(t.leaf_obj, width, data)
        table_content = table_content_for_hash(t.leaf_obj)
        occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Added column {col_idx}"
        )

    if operation == "delete_row":
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        delete_table_row(t.leaf_obj, row)
        table_content = table_content_for_hash(t.leaf_obj)
        occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Deleted row {row}"
        )

    if operation == "delete_column":
        t = resolve_target(doc, target_id)
        if t.leaf_kind != "table":
            raise ValueError("target_id must be a table")
        delete_table_column(t.leaf_obj, col)
        table_content = table_content_for_hash(t.leaf_obj)
        occurrence = count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = f"table_{content_hash(table_content)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Deleted column {col}"
        )

    if operation == "add_page_break":
        p = add_page_break(doc)
        text = p.text or ""
        occurrence = count_occurrence(doc, "paragraph", text, p._element)
        element_id = f"paragraph_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message="Added page break"
        )

    if operation == "add_break":
        t = resolve_target(doc, target_id)
        break_type = content_data or "page"  # default to page break
        p = add_break_after(doc, t.leaf_el, break_type)
        text = p.text or ""
        occurrence = count_occurrence(doc, "paragraph", text, p._element)
        element_id = f"paragraph_{content_hash(text)}_{occurrence}"
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Added {break_type} break after {target_id}",
        )

    if operation == "set_meta":
        meta_data = json.loads(content_data) if content_data else {}
        set_document_meta(
            doc,
            title=meta_data.get("title"),
            author=meta_data.get("author"),
            subject=meta_data.get("subject"),
            keywords=meta_data.get("keywords"),
            category=meta_data.get("category"),
        )
        doc.save(file_path)
        return EditResult(success=True, message="Updated document metadata")

    if operation == "add_section":
        start_type = content_data or "new_page"
        section_idx = add_section(doc, start_type)
        doc.save(file_path)
        return EditResult(
            success=True, message=f"Added section {section_idx} ({start_type})"
        )

    raise ValueError(f"Unknown operation: {operation}")


def main():
    """Entry point for mcp-word command."""
    mcp.run()


if __name__ == "__main__":
    main()
