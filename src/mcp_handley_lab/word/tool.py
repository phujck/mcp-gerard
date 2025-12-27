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
    collect_warnings,
    compute_version,
    create_block,
    delete_block,
    find_body_index_of_element,
    get_document_meta,
    insert_heading_relative,
    insert_paragraph_relative,
    insert_table_relative,
    make_block_id,
    paragraph_kind_and_level,
    replace_paragraph_text,
    replace_table,
    resolve_target,
    table_to_markdown,
)
from mcp_handley_lab.word.models import DocumentReadResult, EditResult

mcp = FastMCP("Word Document Tool")


@mcp.tool(
    description="Read Word document content. Scopes: 'meta' (doc info), 'outline' (headings only), 'blocks' (all content), 'search' (find text)."
)
def read(
    file_path: str = Field(..., description="Path to .docx file"),
    scope: str = Field(
        "outline", description="What to read: 'meta', 'outline', 'blocks', 'search'"
    ),
    search_query: str = Field(
        "", description="Text to search for (required if scope='search')"
    ),
    limit: int = Field(50, description="Max blocks to return"),
    offset: int = Field(0, description="Pagination offset"),
) -> DocumentReadResult:
    """Read Word document content with progressive disclosure."""
    version = compute_version(file_path)
    doc = Document(file_path)
    warnings = collect_warnings(file_path, doc)

    if scope == "meta":
        meta = get_document_meta(doc)
        blocks, block_count = build_blocks(doc, offset=0, limit=0)
        return DocumentReadResult(
            version=version,
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
            version=version,
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    elif scope == "blocks":
        blocks, block_count = build_blocks(doc, offset=offset, limit=limit)
        return DocumentReadResult(
            version=version,
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    elif scope == "search":
        if not search_query:
            return DocumentReadResult(
                version=version,
                block_count=0,
                blocks=[],
                warnings=["search_query is required for scope='search'"],
            )
        blocks, block_count = build_blocks(
            doc, offset=offset, limit=limit, search_query=search_query
        )
        return DocumentReadResult(
            version=version,
            block_count=block_count,
            blocks=blocks,
            warnings=warnings,
        )
    else:
        return DocumentReadResult(
            version=version,
            block_count=0,
            blocks=[],
            warnings=[
                f"Unknown scope: {scope}. Use 'meta', 'outline', 'blocks', or 'search'."
            ],
        )


@mcp.tool(
    description="Edit Word document. Operations: 'create' (new doc), 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style'. All except 'create' require expected_version from read()."
)
def edit(
    file_path: str = Field(..., description="Path to .docx file"),
    expected_version: str = Field(
        "",
        description="Version from read() - required for all operations except 'create'",
    ),
    operation: str = Field(
        ...,
        description="Operation: 'create', 'insert_before', 'insert_after', 'append', 'delete', 'replace', 'style'",
    ),
    target_id: str = Field(
        "",
        description="Block ID from read() - required for insert/delete/replace/style",
    ),
    content_type: str = Field(
        "paragraph",
        description="Type: 'paragraph', 'heading', 'table', 'image', 'list', 'page_break'",
    ),
    content_data: str = Field(
        "", description="Content: text, file path (images), or JSON (tables/lists)"
    ),
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
) -> EditResult:
    """Edit Word document with version check and DOM-style operations."""
    # Handle create operation separately (no version check needed)
    if operation == "create":
        if os.path.exists(file_path):
            return EditResult(
                success=False,
                new_version="",
                message=f"File already exists: {file_path}. Use 'append' to add content to existing documents.",
            )
        doc = Document()
        warnings: list[str] = []
        element_id = ""

        # Optionally add initial content
        if content_data:
            try:
                kind, obj, el = create_block(
                    doc, content_type, content_data, heading_level, style_name
                )
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False,
                    new_version="",
                    message=f"Invalid JSON for {content_type} content_data: {e}",
                )
            except ValueError as e:
                return EditResult(
                    success=False,
                    new_version="",
                    message=str(e),
                )
            except Exception as e:
                return EditResult(
                    success=False,
                    new_version="",
                    message=f"Failed to create {content_type}: {e}",
                )
            idx = find_body_index_of_element(doc, el)
            if kind == "table":
                md, _, _ = table_to_markdown(obj)
                element_id = make_block_id("table", idx or 0, md)
            else:
                block_type, _ = (
                    paragraph_kind_and_level(obj)
                    if kind == "paragraph"
                    else ("paragraph", 0)
                )
                element_id = make_block_id(block_type, idx or 0, obj.text or "")

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Created new document: {file_path}",
            warnings=warnings,
        )

    # All other operations require version check
    current_version = compute_version(file_path)
    if current_version != expected_version:
        return EditResult(
            success=False,
            new_version=current_version,
            message="Document modified since read. Please re-read to get current version.",
        )

    doc = Document(file_path)
    warnings: list[str] = []
    element_id = ""

    if operation in ("insert_before", "insert_after"):
        if not target_id:
            return EditResult(
                success=False,
                new_version=current_version,
                message=f"target_id is required for {operation}",
            )
        try:
            kind, obj, target_el, body_idx = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(
                success=False, new_version=current_version, message=str(e)
            )

        position = "before" if operation == "insert_before" else "after"

        if content_type == "paragraph":
            new_p = insert_paragraph_relative(
                doc, target_el, content_data, position, style_name
            )
            actual_idx = find_body_index_of_element(doc, new_p._element) or 0
            element_id = make_block_id("paragraph", actual_idx, content_data)
        elif content_type == "heading":
            new_p = insert_heading_relative(
                doc, target_el, content_data, heading_level, position
            )
            actual_idx = find_body_index_of_element(doc, new_p._element) or 0
            element_id = make_block_id("heading", actual_idx, content_data)
        elif content_type == "table":
            try:
                table_data = json.loads(content_data)
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False,
                    new_version=current_version,
                    message=f"Invalid JSON for table content_data: {e}",
                )
            new_tbl = insert_table_relative(
                doc, target_el, table_data, position, style_name or "Table Grid"
            )
            actual_idx = find_body_index_of_element(doc, new_tbl._tbl) or 0
            md, _, _ = table_to_markdown(new_tbl)
            element_id = make_block_id("table", actual_idx, md)
        else:
            return EditResult(
                success=False,
                new_version=current_version,
                message=f"content_type '{content_type}' not supported for {operation}. Use 'paragraph', 'heading', or 'table'.",
            )

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Inserted {content_type} {position} {target_id}",
            warnings=warnings,
        )

    elif operation == "append":
        try:
            kind, obj, el = create_block(
                doc, content_type, content_data, heading_level, style_name
            )
        except json.JSONDecodeError as e:
            return EditResult(
                success=False,
                new_version=current_version,
                message=f"Invalid JSON for {content_type} content_data: {e}",
            )
        except ValueError as e:
            return EditResult(
                success=False,
                new_version=current_version,
                message=str(e),
            )
        except Exception as e:
            return EditResult(
                success=False,
                new_version=current_version,
                message=f"Failed to create {content_type}: {e}",
            )
        idx = find_body_index_of_element(doc, el)
        if kind == "table":
            md, _, _ = table_to_markdown(obj)
            element_id = make_block_id("table", idx or 0, md)
        else:
            block_type, _ = (
                paragraph_kind_and_level(obj)
                if kind == "paragraph"
                else ("paragraph", 0)
            )
            element_id = make_block_id(block_type, idx or 0, obj.text or "")

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Appended {content_type} to document",
            warnings=warnings,
        )

    elif operation == "delete":
        if not target_id:
            return EditResult(
                success=False,
                new_version=current_version,
                message="target_id is required for delete",
            )
        try:
            kind, obj, target_el, body_idx = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(
                success=False, new_version=current_version, message=str(e)
            )

        delete_block(kind, obj)
        element_id = target_id

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Deleted block {target_id}",
            warnings=warnings,
        )

    elif operation == "replace":
        if not target_id:
            return EditResult(
                success=False,
                new_version=current_version,
                message="target_id is required for replace",
            )
        try:
            kind, obj, target_el, body_idx = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(
                success=False, new_version=current_version, message=str(e)
            )

        if kind == "paragraph":
            replace_paragraph_text(obj, content_data)
            block_type, _ = paragraph_kind_and_level(obj)
            element_id = make_block_id(block_type, body_idx, content_data)
            warnings.append(
                "Run-level formatting destroyed. Use 'style' operation to reapply formatting."
            )
        elif kind == "table":
            try:
                table_data = json.loads(content_data)
            except json.JSONDecodeError as e:
                return EditResult(
                    success=False,
                    new_version=current_version,
                    message=f"Invalid JSON for table content_data: {e}",
                )
            new_tbl = replace_table(doc, obj, table_data)
            md, _, _ = table_to_markdown(new_tbl)
            element_id = make_block_id("table", body_idx, md)
        else:
            return EditResult(
                success=False,
                new_version=current_version,
                message=f"Cannot replace block of type {kind}",
            )

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Replaced content of {target_id}",
            warnings=warnings,
        )

    elif operation == "style":
        if not target_id:
            return EditResult(
                success=False,
                new_version=current_version,
                message="target_id is required for style",
            )
        try:
            kind, obj, target_el, body_idx = resolve_target(doc, target_id)
        except ValueError as e:
            return EditResult(
                success=False, new_version=current_version, message=str(e)
            )

        if kind != "paragraph":
            if style_name:
                try:
                    obj.style = style_name
                except KeyError:
                    warnings.append(f"Style '{style_name}' not found for table")
            md, _, _ = table_to_markdown(obj)
            element_id = make_block_id("table", body_idx, md)
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
                        success=False,
                        new_version=current_version,
                        message=f"Invalid JSON for formatting: {e}",
                    )
                apply_paragraph_formatting(obj, fmt)
            block_type, _ = paragraph_kind_and_level(obj)
            element_id = make_block_id(block_type, body_idx, obj.text or "")

        doc.save(file_path)
        new_version = compute_version(file_path)
        return EditResult(
            success=True,
            new_version=new_version,
            element_id=element_id,
            message=f"Applied style to {target_id}",
            warnings=warnings,
        )

    else:
        return EditResult(
            success=False,
            new_version=current_version,
            message=f"Unknown operation: {operation}. Use 'insert_before', 'insert_after', 'append', 'delete', 'replace', or 'style'.",
        )


def main():
    """Run the Word document MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
