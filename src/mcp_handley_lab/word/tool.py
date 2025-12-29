"""Word document MCP tool - read and edit operations."""

import json

from docx import Document
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.word import document as word_ops
from mcp_handley_lab.word.models import DocumentReadResult, EditResult

mcp = FastMCP("Word Document Tool")


@mcp.tool(
    description="Read Word document content. Scopes: 'meta' (doc info), 'outline' (headings only), 'blocks' (all content), 'search' (find text), 'table_cells' (cells of a table), 'table_layout' (table alignment/autofit/row heights), 'runs' (text runs in a paragraph), 'comments' (all comments), 'headers_footers' (headers/footers per section), 'page_setup' (margins, orientation per section), 'images' (embedded inline images), 'hyperlinks' (all hyperlinks with URLs), 'styles' (all document styles), 'style' (detailed formatting for a specific style by name in target_id). Block IDs are content-addressed (type_hash_occurrence) and stable across structural edits."
)
def read(
    file_path: str = Field(..., description="Path to .docx file"),
    scope: str = Field(
        "outline",
        description="What to read: 'meta', 'outline', 'blocks', 'search', 'table_cells', 'table_layout', 'runs', 'comments', 'headers_footers', 'page_setup', 'images', 'hyperlinks', 'styles', 'style'",
    ),
    target_id: str = Field(
        "",
        description="Block ID for table_cells/runs/table_layout scopes, or style name for 'style' scope",
    ),
    search_query: str = Field("", description="Text to search for (scope='search')"),
    limit: int = Field(50, description="Max blocks to return"),
    offset: int = Field(0, description="Pagination offset"),
) -> DocumentReadResult:
    """Read Word document content."""
    doc = Document(file_path)

    if scope == "meta":
        meta = word_ops.get_document_meta(doc)
        _, block_count = word_ops.build_blocks(doc, offset=0, limit=0)
        return DocumentReadResult(block_count=block_count, meta=meta)
    if scope == "outline":
        blocks, block_count = word_ops.build_blocks(
            doc, offset=offset, limit=limit, heading_only=True
        )
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "blocks":
        blocks, block_count = word_ops.build_blocks(doc, offset=offset, limit=limit)
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "search":
        blocks, block_count = word_ops.build_blocks(
            doc, offset=offset, limit=limit, search_query=search_query
        )
        return DocumentReadResult(block_count=block_count, blocks=blocks)
    if scope == "table_cells":
        t = word_ops.resolve_target(doc, target_id)
        cells = word_ops.build_table_cells(t.leaf_obj, t.base_id)
        return DocumentReadResult(
            block_count=len(cells),
            cells=cells,
            table_rows=len(t.leaf_obj.rows),
            table_cols=len(t.leaf_obj.columns),
        )
    if scope == "table_layout":
        t = word_ops.resolve_target(doc, target_id)
        table_layout = word_ops.build_table_layout(t.leaf_obj, t.base_id)
        return DocumentReadResult(block_count=1, table_layout=table_layout)
    if scope == "runs":
        t = word_ops.resolve_target(doc, target_id)
        runs = word_ops.build_runs(t.leaf_obj)
        paragraph_format = word_ops.build_paragraph_format(t.leaf_obj)
        return DocumentReadResult(
            block_count=len(runs), runs=runs, paragraph_format=paragraph_format
        )
    if scope == "comments":
        comments = word_ops.build_comments(doc)
        return DocumentReadResult(block_count=len(comments), comments=comments)
    if scope == "headers_footers":
        headers_footers = word_ops.build_headers_footers(doc)
        return DocumentReadResult(
            block_count=len(headers_footers), headers_footers=headers_footers
        )
    if scope == "page_setup":
        page_setup = word_ops.build_page_setup(doc)
        return DocumentReadResult(block_count=len(page_setup), page_setup=page_setup)
    if scope == "images":
        images = word_ops.build_images(doc)
        return DocumentReadResult(block_count=len(images), images=images)
    if scope == "hyperlinks":
        hyperlinks = word_ops.build_hyperlinks(doc)
        return DocumentReadResult(block_count=len(hyperlinks), hyperlinks=hyperlinks)
    if scope == "styles":
        styles = word_ops.build_styles(doc)
        return DocumentReadResult(block_count=len(styles), styles=styles)
    if scope == "style":
        style_format = word_ops.get_style_format(doc, target_id)
        return DocumentReadResult(block_count=1, style_format=style_format)
    raise ValueError(f"Unknown scope: {scope}")


@mcp.tool(
    description="Edit Word document. Operations: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'edit_style', 'add_comment', 'set_header', 'set_footer', 'set_first_page_header', 'set_first_page_footer', 'set_even_page_header', 'set_even_page_footer', 'append_header', 'append_footer', 'clear_header', 'clear_footer', 'set_margins', 'set_orientation', 'insert_image', 'delete_image', 'add_row', 'add_column', 'delete_row', 'delete_column', 'add_page_break', 'add_break', 'set_meta', 'add_section', 'merge_cells', 'set_table_alignment', 'set_table_autofit', 'set_table_fixed_layout', 'set_row_height', 'set_cell_width', 'set_cell_vertical_alignment', 'add_tab_stop', 'clear_tab_stops', 'insert_field', 'insert_page_x_of_y'. Block IDs are content-addressed and stable across structural edits. Returns new ID after content changes."
)
def edit(
    file_path: str = Field(..., description="Path to .docx file"),
    operation: str = Field(
        ...,
        description="Operation: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run', 'edit_style', 'add_comment', 'set_header', 'set_footer', 'set_first_page_header', 'set_first_page_footer', 'set_even_page_header', 'set_even_page_footer', 'append_header', 'append_footer', 'clear_header', 'clear_footer', 'set_margins', 'set_orientation', 'insert_image', 'delete_image', 'add_row', 'add_column', 'delete_row', 'delete_column', 'add_page_break', 'add_break', 'set_meta', 'add_section', 'merge_cells', 'set_table_alignment', 'set_table_autofit', 'set_table_fixed_layout', 'set_row_height', 'set_cell_width', 'set_cell_vertical_alignment', 'add_tab_stop', 'clear_tab_stops', 'insert_field', 'insert_page_x_of_y'",
    ),
    target_id: str = Field(
        "",
        description="Block ID from read() - required for insert/delete/replace/style/edit_cell/edit_run/add_comment/insert_image and table operations. For edit_style: style name. For delete_image: the image ID.",
    ),
    content_type: str = Field(
        "paragraph",
        description="Type: 'paragraph', 'heading', 'table'",
    ),
    content_data: str = Field(
        "",
        description="Content: text or JSON. For set_table_alignment: 'left'/'center'/'right'. For set_table_autofit: 'true'/'false'. For set_table_fixed_layout: JSON array of widths. For set_row_height: JSON {height, rule}. For set_cell_width: width in inches. For set_cell_vertical_alignment: 'top'/'center'/'bottom'. For add_tab_stop: JSON {position, alignment, leader}. For insert_field: field code (PAGE, NUMPAGES, DATE, TIME). For insert_page_x_of_y: 'header' or 'footer'.",
    ),
    style_name: str = Field(
        "", description="Apply Word style: 'Heading 1', 'Normal', etc."
    ),
    formatting: str = Field(
        "",
        description='Direct formatting JSON. Text/Run: {"bold": true, "color": "FF0000", "font_size": 14, "highlight_color": "yellow", "strike": true, "subscript": true, "superscript": true, "style": "Strong"}. Character styles: "Strong", "Emphasis", "Hyperlink", etc. Paragraph: {"left_indent": 0.5, "right_indent": 0.5, "first_line_indent": 0.5, "space_before": 12, "space_after": 12, "line_spacing": 1.5, "keep_with_next": true, "page_break_before": true} (indents in inches, spacing in points, line_spacing < 5 is multiplier). Margins: {"top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25} in inches. Images: {"width": 4} or {"height": 3} in inches.',
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
    # Create operation is special - creates new document
    if operation == "create":
        doc = Document()
        obj = word_ops.create_element(
            doc, content_type, content_data, style_name, heading_level
        )
        level = heading_level if content_type == "heading" else 0
        element_id = word_ops.get_element_id(doc, obj, level)
        doc.save(file_path)
        return EditResult(
            success=True, element_id=element_id, message=f"Created: {file_path}"
        )

    # All other operations work on existing document
    doc = Document(file_path)
    element_id = target_id
    comment_id = None
    message = f"Completed {operation}"

    if operation in ("insert_before", "insert_after"):
        t = word_ops.resolve_target(doc, target_id)
        position = "before" if operation == "insert_before" else "after"
        obj = word_ops.insert_content_relative(
            doc,
            t.leaf_el,
            position,
            content_type,
            content_data,
            style_name,
            heading_level,
        )
        level = heading_level if content_type == "heading" else 0
        element_id = word_ops.get_element_id(doc, obj, level)
        message = f"Inserted {content_type} {position} {target_id}"

    elif operation == "append":
        obj = word_ops.create_element(
            doc, content_type, content_data, style_name, heading_level
        )
        level = heading_level if content_type == "heading" else 0
        element_id = word_ops.get_element_id(doc, obj, level)
        message = f"Appended {content_type} to document"

    elif operation == "delete":
        t = word_ops.resolve_target(doc, target_id)
        word_ops.delete_block(t.base_obj)
        message = f"Deleted block {target_id}"

    elif operation == "replace":
        t = word_ops.resolve_target(doc, target_id)
        if t.leaf_kind.startswith("heading") or t.leaf_kind == "paragraph":
            t.leaf_obj.text = content_data
            text = content_data
            occurrence = word_ops.count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
            element_id = word_ops.make_block_id(t.leaf_kind, text, occurrence)
        elif t.leaf_kind == "table":
            table_data = json.loads(content_data)
            new_tbl = word_ops.replace_table(doc, t.leaf_obj, table_data)
            table_content = word_ops.table_content_for_hash(new_tbl)
            occurrence = word_ops.count_occurrence(
                doc, "table", table_content, new_tbl._tbl
            )
            element_id = word_ops.make_block_id("table", table_content, occurrence)
        else:
            raise ValueError(f"Unsupported leaf_kind: {t.leaf_kind}")
        message = f"Replaced content of {target_id}"

    elif operation == "style":
        t = word_ops.resolve_target(doc, target_id)
        if style_name:
            t.leaf_obj.style = style_name
        if t.leaf_kind == "table":
            table_content = word_ops.table_content_for_hash(t.leaf_obj)
            occurrence = word_ops.count_occurrence(
                doc, "table", table_content, t.leaf_el
            )
            element_id = word_ops.make_block_id("table", table_content, occurrence)
        else:
            if formatting:
                fmt = json.loads(formatting)
                word_ops.apply_paragraph_formatting(t.leaf_obj, fmt)
            text = t.leaf_obj.text or ""
            occurrence = word_ops.count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
            element_id = word_ops.make_block_id(t.leaf_kind, text, occurrence)
        message = f"Applied style to {target_id}"

    elif operation == "edit_cell":
        t = word_ops.resolve_target(doc, target_id)
        word_ops.replace_table_cell(t.leaf_obj, row, col, content_data)
        table_content = word_ops.table_content_for_hash(t.leaf_obj)
        occurrence = word_ops.count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = word_ops.make_block_id("table", table_content, occurrence)
        message = f"Updated cell r{row}c{col}"

    elif operation == "edit_run":
        t = word_ops.resolve_target(doc, target_id)
        if content_data:
            word_ops.edit_run_text(t.leaf_obj, run_index, content_data)
        if formatting:
            fmt = json.loads(formatting)
            word_ops.edit_run_formatting(t.leaf_obj, run_index, fmt)
        text = t.leaf_obj.text or ""
        occurrence = word_ops.count_occurrence(doc, t.leaf_kind, text, t.leaf_el)
        element_id = word_ops.make_block_id(t.leaf_kind, text, occurrence)
        message = f"Updated run {run_index}"

    elif operation == "edit_style":
        fmt = json.loads(formatting)
        word_ops.edit_style(doc, target_id, fmt)
        message = f"Modified style: {target_id}"

    elif operation == "add_comment":
        t = word_ops.resolve_target(doc, target_id)
        comment_id = word_ops.add_comment_to_block(
            doc, t.leaf_obj, content_data, author, initials
        )
        message = f"Added comment {comment_id} to {target_id}"

    elif operation in (
        "set_header",
        "set_footer",
        "set_first_page_header",
        "set_first_page_footer",
        "set_even_page_header",
        "set_even_page_footer",
    ):
        location = operation[4:]  # Remove "set_" prefix
        word_ops.set_header_footer_text(doc, section_index, content_data, location)
        message = f"Set {location.replace('_', ' ')} for section {section_index}"

    elif operation in ("append_header", "append_footer"):
        location = operation[7:]  # Remove "append_" prefix
        element_id = word_ops.append_to_header_footer(
            doc, section_index, content_type, content_data, location
        )
        message = f"Appended {content_type} to {location} of section {section_index}"

    elif operation in ("clear_header", "clear_footer"):
        location = operation[6:]  # Remove "clear_" prefix
        word_ops.clear_header_footer(doc, section_index, location)
        message = f"Cleared {location} for section {section_index}"

    elif operation == "set_margins":
        margins = json.loads(formatting)
        word_ops.set_page_margins(
            doc,
            section_index,
            top=margins["top"],
            bottom=margins["bottom"],
            left=margins["left"],
            right=margins["right"],
        )
        message = f"Set margins for section {section_index}"

    elif operation == "set_orientation":
        word_ops.set_page_orientation(doc, section_index, content_data)
        message = f"Set orientation to {content_data} for section {section_index}"

    elif operation == "insert_image":
        fmt = json.loads(formatting) if formatting else {}
        element_id = word_ops.insert_image(
            doc,
            content_data,
            target_id,
            "after",
            width_inches=float(fmt.get("width", 0)),
            height_inches=float(fmt.get("height", 0)),
        )
        message = "Inserted image"

    elif operation == "delete_image":
        word_ops.delete_image(doc, target_id)
        message = f"Deleted {target_id}"

    elif operation == "add_row":
        t = word_ops.resolve_target(doc, target_id)
        data = json.loads(content_data) if content_data else None
        row_idx = word_ops.add_table_row(t.leaf_obj, data)
        table_content = word_ops.table_content_for_hash(t.leaf_obj)
        occurrence = word_ops.count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = word_ops.make_block_id("table", table_content, occurrence)
        message = f"Added row {row_idx}"

    elif operation == "add_column":
        t = word_ops.resolve_target(doc, target_id)
        fmt = json.loads(formatting) if formatting else {}
        width = float(fmt.get("width", 1.0))
        data = json.loads(content_data) if content_data else None
        col_idx = word_ops.add_table_column(t.leaf_obj, width, data)
        table_content = word_ops.table_content_for_hash(t.leaf_obj)
        occurrence = word_ops.count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = word_ops.make_block_id("table", table_content, occurrence)
        message = f"Added column {col_idx}"

    elif operation == "delete_row":
        t = word_ops.resolve_target(doc, target_id)
        word_ops.delete_table_row(t.leaf_obj, row)
        table_content = word_ops.table_content_for_hash(t.leaf_obj)
        occurrence = word_ops.count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = word_ops.make_block_id("table", table_content, occurrence)
        message = f"Deleted row {row}"

    elif operation == "delete_column":
        t = word_ops.resolve_target(doc, target_id)
        word_ops.delete_table_column(t.leaf_obj, col)
        table_content = word_ops.table_content_for_hash(t.leaf_obj)
        occurrence = word_ops.count_occurrence(doc, "table", table_content, t.leaf_el)
        element_id = word_ops.make_block_id("table", table_content, occurrence)
        message = f"Deleted column {col}"

    elif operation == "add_page_break":
        p = word_ops.add_page_break(doc)
        text = p.text or ""
        occurrence = word_ops.count_occurrence(doc, "paragraph", text, p._element)
        element_id = word_ops.make_block_id("paragraph", text, occurrence)
        message = "Added page break"

    elif operation == "add_break":
        t = word_ops.resolve_target(doc, target_id)
        break_type = content_data or "page"
        p = word_ops.add_break_after(doc, t.leaf_el, break_type)
        text = p.text or ""
        occurrence = word_ops.count_occurrence(doc, "paragraph", text, p._element)
        element_id = word_ops.make_block_id("paragraph", text, occurrence)
        message = f"Added {break_type} break after {target_id}"

    elif operation == "set_meta":
        meta_data = json.loads(content_data) if content_data else {}
        word_ops.set_document_meta(
            doc,
            title=meta_data.get("title"),
            author=meta_data.get("author"),
            subject=meta_data.get("subject"),
            keywords=meta_data.get("keywords"),
            category=meta_data.get("category"),
        )
        message = "Updated document metadata"

    elif operation == "add_section":
        start_type = content_data or "new_page"
        section_idx = word_ops.add_section(doc, start_type)
        message = f"Added section {section_idx} ({start_type})"

    elif operation == "merge_cells":
        target = word_ops.resolve_target(doc, target_id)
        end_data = json.loads(content_data)
        word_ops.merge_cells(
            target.base_obj, row, col, end_data["end_row"], end_data["end_col"]
        )
        message = f"Merged cells from ({row},{col}) to ({end_data['end_row']},{end_data['end_col']})"

    elif operation == "set_table_alignment":
        target = word_ops.resolve_target(doc, target_id)
        word_ops.set_table_alignment(target.base_obj, content_data)
        message = f"Set table alignment to {content_data}"

    elif operation == "set_table_autofit":
        target = word_ops.resolve_target(doc, target_id)
        target.base_obj.autofit = json.loads(content_data.lower())
        message = f"Set table autofit to {target.base_obj.autofit}"

    elif operation == "set_table_fixed_layout":
        target = word_ops.resolve_target(doc, target_id)
        widths = json.loads(content_data)
        word_ops.set_table_fixed_layout(target.base_obj, widths)
        message = f"Set table fixed layout with {len(widths)} columns"

    elif operation == "set_row_height":
        target = word_ops.resolve_target(doc, target_id)
        height_data = json.loads(content_data)
        word_ops.set_row_height(
            target.base_obj,
            row,
            height_data["height"],
            height_data.get("rule", "at_least"),
        )
        message = f"Set row {row} height to {height_data['height']} inches"

    elif operation == "set_cell_width":
        target = word_ops.resolve_target(doc, target_id)
        word_ops.set_cell_width(target.base_obj, row, col, float(content_data))
        message = f"Set cell ({row},{col}) width to {content_data} inches"

    elif operation == "set_cell_vertical_alignment":
        target = word_ops.resolve_target(doc, target_id)
        word_ops.set_cell_vertical_alignment(target.base_obj, row, col, content_data)
        message = f"Set cell ({row},{col}) vertical alignment to {content_data}"

    elif operation == "add_tab_stop":
        target = word_ops.resolve_target(doc, target_id)
        para = (
            target.leaf_obj.paragraphs[0]
            if target.leaf_kind == "cell"
            else target.leaf_obj
        )
        tab_data = json.loads(content_data)
        word_ops.add_tab_stop(
            para,
            float(tab_data["position"]),
            tab_data.get("alignment", "left"),
            tab_data.get("leader", "spaces"),
        )
        message = f"Added tab stop at {tab_data['position']} inches ({tab_data.get('alignment', 'left')}, {tab_data.get('leader', 'spaces')})"

    elif operation == "clear_tab_stops":
        target = word_ops.resolve_target(doc, target_id)
        para = (
            target.leaf_obj.paragraphs[0]
            if target.leaf_kind == "cell"
            else target.leaf_obj
        )
        para.paragraph_format.tab_stops.clear_all()
        message = "Cleared all tab stops"

    elif operation == "insert_field":
        target = word_ops.resolve_target(doc, target_id)
        para = (
            target.leaf_obj.paragraphs[0]
            if target.leaf_kind == "cell"
            else target.leaf_obj
        )
        field_code = content_data.strip().upper()
        word_ops.insert_field(para, field_code)
        message = f"Inserted {field_code} field"

    elif operation == "insert_page_x_of_y":
        location = content_data.strip().lower() or "footer"
        word_ops.insert_page_x_of_y(doc, section_index, location)
        message = f"Inserted 'Page X of Y' in {location} of section {section_index}"

    else:
        raise ValueError(f"Unknown operation: {operation}")

    doc.save(file_path)
    return EditResult(
        success=True, element_id=element_id, comment_id=comment_id, message=message
    )


if __name__ == "__main__":
    mcp.run()
