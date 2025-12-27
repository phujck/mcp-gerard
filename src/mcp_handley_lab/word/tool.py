"""Word document MCP tool - read and edit operations."""

import json
import os

from docx import Document
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.word.document import (
    apply_paragraph_formatting,
    apply_paragraph_style,
    build_blocks,
    build_runs,
    build_table_cells,
    collect_warnings,
    count_occurrence,
    create_block,
    delete_block,
    edit_run_formatting,
    edit_run_text,
    get_document_meta,
    insert_heading_relative,
    insert_paragraph_relative,
    insert_table_relative,
    make_block_id,
    replace_paragraph_text,
    replace_table,
    replace_table_cell,
    resolve_target,
    table_content_for_hash,
)
from mcp_handley_lab.word.models import DocumentReadResult, EditResult

mcp = FastMCP("Word Document Tool")


@mcp.tool(
    description="Read Word document content. Scopes: 'meta' (doc info), 'outline' (headings only), 'blocks' (all content), 'search' (find text), 'table_cells' (cells of a table), 'runs' (text runs in a paragraph). Block IDs are content-addressed (type_hash_occurrence) and stable across structural edits."
)
def read(
    file_path: str = Field(..., description="Path to .docx file"),
    scope: str = Field(
        "outline",
        description="What to read: 'meta', 'outline', 'blocks', 'search', 'table_cells', 'runs'",
    ),
    target_id: str = Field(
        "", description="Block ID (required for scope='table_cells' or 'runs')"
    ),
    search_query: str = Field(
        "", description="Text to search for (required if scope='search')"
    ),
    limit: int = Field(50, description="Max blocks to return"),
    offset: int = Field(0, description="Pagination offset"),
) -> DocumentReadResult:
    """Read Word document content with progressive disclosure."""
    doc = Document(file_path)
    warnings = collect_warnings(file_path)

    if scope == "meta":
        meta = get_document_meta(doc)
        blocks, block_count = build_blocks(doc, offset=0, limit=0)
        return DocumentReadResult(
            block_count=block_count,
            blocks=[],
            meta=meta,
            warnings=warnings,
        )
    elif scope == "outline":
        blocks, block_count = build_blocks(
            doc, offset=offset, limit=limit, heading_only=True
        )
        return DocumentReadResult(
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    elif scope == "blocks":
        blocks, block_count = build_blocks(doc, offset=offset, limit=limit)
        return DocumentReadResult(
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    elif scope == "search":
        if not search_query:
            return DocumentReadResult(
                block_count=0,
                blocks=[],
                warnings=["search_query is required for scope='search'"],
            )
        blocks, block_count = build_blocks(
            doc, offset=offset, limit=limit, search_query=search_query
        )
        return DocumentReadResult(
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    elif scope == "table_cells":
        if not target_id:
            return DocumentReadResult(
                block_count=0,
                warnings=["target_id is required for scope='table_cells'"],
            )
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return DocumentReadResult(
                block_count=0,
                warnings=[str(e)],
            )
        if block_type != "table":
            return DocumentReadResult(
                block_count=0,
                warnings=[f"target_id must be a table, got {block_type}"],
            )
        cells = build_table_cells(obj)
        return DocumentReadResult(
            block_count=len(cells),
            cells=cells,
            table_rows=len(obj.rows),
            table_cols=len(obj.columns),
            warnings=warnings,
        )
    elif scope == "runs":
        if not target_id:
            return DocumentReadResult(
                block_count=0,
                warnings=["target_id is required for scope='runs'"],
            )
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return DocumentReadResult(
                block_count=0,
                warnings=[str(e)],
            )
        if block_type == "table":
            return DocumentReadResult(
                block_count=0,
                warnings=["scope='runs' requires a paragraph or heading, not a table"],
            )
        runs, run_warnings = build_runs(obj)
        warnings.extend(run_warnings)
        return DocumentReadResult(
            block_count=len(runs),
            runs=runs,
            warnings=warnings,
        )
    else:
        return DocumentReadResult(
            block_count=0,
            blocks=[],
            warnings=[
                f"Unknown scope: {scope}. Use 'meta', 'outline', 'blocks', 'search', 'table_cells', or 'runs'."
            ],
        )


@mcp.tool(
    description="Edit Word document. Operations: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run'. Block IDs are content-addressed and stable across structural edits. Returns new ID after content changes."
)
def edit(
    file_path: str = Field(..., description="Path to .docx file"),
    operation: str = Field(
        ...,
        description="Operation: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', 'edit_run'",
    ),
    target_id: str = Field(
        "",
        description="Block ID from read() - required for insert/delete/replace/style/edit_cell/edit_run",
    ),
    content_type: str = Field(
        "paragraph",
        description="Type: 'paragraph', 'heading', 'table'",
    ),
    content_data: str = Field("", description="Content: text or JSON (for tables)"),
    style_name: str = Field(
        "", description="Apply Word style: 'Heading 1', 'Normal', etc."
    ),
    formatting: str = Field(
        "",
        description='Direct formatting JSON: {"bold": true, "color": "FF0000", "font_size": 14}',
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
) -> EditResult:
    """Edit Word document."""
    if operation == "create":
        if os.path.exists(file_path):
            return EditResult(
                success=False,
                message=f"File already exists: {file_path}. Use 'append' to add content.",
            )
        doc = Document()
        element_id = ""

        if content_data:
            try:
                block_type, obj, el = create_block(
                    doc, content_type, content_data, heading_level, style_name
                )
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False,
                    message=f"Invalid JSON for {content_type}: {e}",
                )
            except ValueError as e:
                return EditResult(success=False, message=str(e))
            # Compute content-hash ID with occurrence
            if block_type == "table":
                text = table_content_for_hash(obj)
            else:
                text = obj.text or ""
            occurrence = count_occurrence(doc, block_type, text, el)
            element_id = make_block_id(block_type, text, occurrence)

        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Created: {file_path}",
        )

    doc = Document(file_path)
    element_id = ""

    if operation in ("insert_before", "insert_after"):
        if not target_id:
            return EditResult(
                success=False, message=f"target_id required for {operation}"
            )
        try:
            _, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))

        position = "before" if operation == "insert_before" else "after"

        if content_type == "paragraph":
            new_p = insert_paragraph_relative(
                doc, target_el, content_data, position, style_name
            )
            text = new_p.text or ""
            occurrence = count_occurrence(doc, "paragraph", text, new_p._element)
            element_id = make_block_id("paragraph", text, occurrence)
        elif content_type == "heading":
            new_p = insert_heading_relative(
                doc, target_el, content_data, heading_level, position
            )
            text = new_p.text or ""
            block_type = f"heading{max(1, min(heading_level, 9))}"
            occurrence = count_occurrence(doc, block_type, text, new_p._element)
            element_id = make_block_id(block_type, text, occurrence)
        elif content_type == "table":
            if not content_data:
                return EditResult(
                    success=False, message="content_data is required for table"
                )
            try:
                table_data = json.loads(content_data)
            except json.JSONDecodeError as e:
                return EditResult(success=False, message=f"Invalid JSON for table: {e}")
            new_tbl = insert_table_relative(
                doc, target_el, table_data, position, style_name or "Table Grid"
            )
            table_content = table_content_for_hash(new_tbl)
            occurrence = count_occurrence(doc, "table", table_content, new_tbl._tbl)
            element_id = make_block_id("table", table_content, occurrence)
        else:
            return EditResult(
                success=False, message=f"Unsupported content_type: {content_type}"
            )

        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Inserted {content_type} {position} {target_id}",
        )

    elif operation == "append":
        if content_type == "table" and not content_data:
            return EditResult(
                success=False, message="content_data is required for table"
            )
        try:
            block_type, obj, el = create_block(
                doc, content_type, content_data, heading_level, style_name
            )
        except json.JSONDecodeError as e:
            return EditResult(
                success=False,
                message=f"Invalid JSON for {content_type} content_data: {e}",
            )
        except ValueError as e:
            return EditResult(success=False, message=str(e))
        except Exception as e:
            return EditResult(
                success=False,
                message=f"Failed to create {content_type}: {e}",
            )
        # Compute content-hash ID with occurrence
        text = table_content_for_hash(obj) if block_type == "table" else obj.text or ""
        occurrence = count_occurrence(doc, block_type, text, el)
        element_id = make_block_id(block_type, text, occurrence)

        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Appended {content_type} to document",
        )

    elif operation == "delete":
        if not target_id:
            return EditResult(success=False, message="target_id is required for delete")
        try:
            block_type, obj, _, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))

        delete_block(block_type, obj)
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=target_id,
            message=f"Deleted block {target_id}",
        )

    elif operation == "replace":
        if not target_id:
            return EditResult(
                success=False, message="target_id is required for replace"
            )
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))

        if block_type.startswith("heading") or block_type == "paragraph":
            replace_paragraph_text(obj, content_data)
            text = content_data
            occurrence = count_occurrence(doc, block_type, text, target_el)
            element_id = make_block_id(block_type, text, occurrence)
        elif block_type == "table":
            if not content_data:
                return EditResult(
                    success=False, message="content_data is required for table"
                )
            try:
                table_data = json.loads(content_data)
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False,
                    message=f"Invalid JSON for table content_data: {e}",
                )
            new_tbl = replace_table(doc, obj, table_data)
            table_content = table_content_for_hash(new_tbl)
            occurrence = count_occurrence(doc, "table", table_content, new_tbl._tbl)
            element_id = make_block_id("table", table_content, occurrence)
        else:
            return EditResult(
                success=False, message=f"Cannot replace block of type {block_type}"
            )

        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Replaced content of {target_id}",
        )

    elif operation == "style":
        if not target_id:
            return EditResult(success=False, message="target_id is required for style")
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))

        warnings = []
        if block_type == "table":
            if style_name:
                try:
                    obj.style = style_name
                except KeyError:
                    warnings.append(f"Style '{style_name}' not found for table")
            table_content = table_content_for_hash(obj)
            occurrence = count_occurrence(doc, "table", table_content, target_el)
            element_id = make_block_id("table", table_content, occurrence)
        else:
            if style_name:
                try:
                    apply_paragraph_style(obj, style_name)
                except KeyError:
                    warnings.append(f"Style '{style_name}' not found, using 'Normal'")
                    apply_paragraph_style(obj, "Normal")
            if formatting:
                try:
                    fmt = json.loads(formatting)
                except json.JSONDecodeError as e:
                    return EditResult(
                        success=False, message=f"Invalid JSON for formatting: {e}"
                    )
                apply_paragraph_formatting(obj, fmt)
            text = obj.text or ""
            occurrence = count_occurrence(doc, block_type, text, target_el)
            element_id = make_block_id(block_type, text, occurrence)

        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Applied style to {target_id}",
            warnings=warnings if warnings else [],
        )

    elif operation == "edit_cell":
        if not target_id:
            return EditResult(
                success=False, message="target_id is required for edit_cell"
            )
        if row < 1 or col < 1:
            return EditResult(
                success=False, message="row and col must be >= 1 (1-based)"
            )
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))
        if block_type != "table":
            return EditResult(
                success=False, message=f"target_id must be a table, got {block_type}"
            )
        try:
            replace_table_cell(obj, row, col, content_data)
        except ValueError as e:
            return EditResult(success=False, message=str(e))
        # Return updated table block ID (hash changes after cell edit)
        table_content = table_content_for_hash(obj)
        occurrence = count_occurrence(doc, "table", table_content, target_el)
        element_id = make_block_id("table", table_content, occurrence)
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Updated cell r{row}c{col}",
        )

    elif operation == "edit_run":
        if not target_id:
            return EditResult(
                success=False, message="target_id is required for edit_run"
            )
        if run_index < 0:
            return EditResult(
                success=False,
                message="run_index is required for edit_run (0-based). Use read() with scope='runs' to find indices.",
            )
        try:
            block_type, obj, target_el, _ = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(success=False, message=str(e))
        if block_type == "table":
            return EditResult(
                success=False,
                message="edit_run requires a paragraph or heading, not a table",
            )

        warnings = []
        # Edit run text if content_data provided
        if content_data:
            try:
                text_warnings = edit_run_text(obj, run_index, content_data)
                warnings.extend(text_warnings)
            except ValueError as e:
                return EditResult(success=False, message=str(e))

        # Apply formatting if provided
        if formatting:
            try:
                fmt = json.loads(formatting)
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False, message=f"Invalid JSON for formatting: {e}"
                )
            try:
                edit_run_formatting(obj, run_index, fmt)
            except ValueError as e:
                return EditResult(success=False, message=str(e))

        # Return updated block ID (text hash changes after run edit)
        text = obj.text or ""
        occurrence = count_occurrence(doc, block_type, text, target_el)
        element_id = make_block_id(block_type, text, occurrence)
        doc.save(file_path)
        return EditResult(
            success=True,
            element_id=element_id,
            message=f"Updated run {run_index}",
            warnings=warnings,
        )

    else:
        return EditResult(
            success=False,
            message=f"Unknown operation: {operation}. Use 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style', 'edit_cell', or 'edit_run'.",
        )


def main():
    """Run the Word document MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
