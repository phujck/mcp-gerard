"""Integration tests for Word document MCP tool."""

import json
import tempfile
from pathlib import Path

import pytest
from docx import Document

from mcp_handley_lab.word.tool import mcp


@pytest.fixture
def sample_docx():
    """Create a sample Word document for testing."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        doc = Document()
        doc.add_heading("Test Document", level=1)
        doc.add_paragraph("This is the first paragraph.")
        doc.add_heading("Section Two", level=2)
        doc.add_paragraph("This is the second paragraph.")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Header 1"
        table.cell(0, 1).text = "Header 2"
        table.cell(1, 0).text = "Row 1 Col 1"
        table.cell(1, 1).text = "Row 1 Col 2"
        doc.save(f.name)
        yield Path(f.name)
    Path(f.name).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_read_outline(sample_docx):
    """Test reading document outline (headings only)."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "outline"}
    )
    assert result["block_count"] == 2  # Two headings
    assert len(result["blocks"]) == 2
    assert result["blocks"][0]["type"] == "heading1"  # Level in type
    assert result["blocks"][0]["level"] == 1
    assert "Test Document" in result["blocks"][0]["text"]
    assert result["blocks"][1]["type"] == "heading2"
    assert result["blocks"][1]["level"] == 2


@pytest.mark.asyncio
async def test_read_meta(sample_docx):
    """Test reading document metadata."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "meta"}
    )
    assert result["meta"] is not None
    assert result["blocks"] == []


@pytest.mark.asyncio
async def test_read_blocks(sample_docx):
    """Test reading all blocks."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    assert result["block_count"] == 5  # 2 headings + 2 paragraphs + 1 table
    types = [b["type"] for b in result["blocks"]]
    assert "heading1" in types  # Level in type
    assert "heading2" in types
    assert "paragraph" in types
    assert "table" in types


@pytest.mark.asyncio
async def test_read_search(sample_docx):
    """Test searching for text."""
    _, result = await mcp.call_tool(
        "read",
        {"file_path": str(sample_docx), "scope": "search", "search_query": "second"},
    )
    assert result["block_count"] == 1  # Only matching blocks counted
    assert len(result["blocks"]) == 1
    assert "second" in result["blocks"][0]["text"].lower()


@pytest.mark.asyncio
async def test_edit_append_paragraph(sample_docx):
    """Test appending a paragraph."""
    _, read_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    initial_count = read_result["block_count"]

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "append",
            "content_type": "paragraph",
            "content_data": "New paragraph at the end",
        },
    )
    assert edit_result["success"]
    assert edit_result["element_id"]

    _, read_result2 = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    assert read_result2["block_count"] == initial_count + 1


@pytest.mark.asyncio
async def test_edit_delete(sample_docx):
    """Test deleting a block."""
    _, read_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    initial_count = read_result["block_count"]
    target_id = read_result["blocks"][1]["id"]  # Delete second block

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "delete",
            "target_id": target_id,
        },
    )
    assert edit_result["success"]

    _, read_result2 = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    assert read_result2["block_count"] == initial_count - 1


@pytest.mark.asyncio
async def test_edit_insert_before(sample_docx):
    """Test inserting before a block."""
    _, read_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    target_id = read_result["blocks"][1]["id"]

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "insert_before",
            "target_id": target_id,
            "content_type": "paragraph",
            "content_data": "Inserted before",
        },
    )
    assert edit_result["success"]
    assert edit_result["element_id"]


@pytest.mark.asyncio
async def test_edit_replace(sample_docx):
    """Test replacing block content."""
    _, read_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    para_block = next(b for b in read_result["blocks"] if b["type"] == "paragraph")

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "replace",
            "target_id": para_block["id"],
            "content_data": "Replaced content",
        },
    )
    assert edit_result["success"]


@pytest.mark.asyncio
async def test_edit_append_table(sample_docx):
    """Test appending a table."""
    table_data = json.dumps([["A", "B"], ["1", "2"]])
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "append",
            "content_type": "table",
            "content_data": table_data,
        },
    )
    assert edit_result["success"]
    assert edit_result["element_id"]


@pytest.mark.asyncio
async def test_edit_create_empty():
    """Test creating a new empty document."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        new_path = Path(f.name)
    new_path.unlink()  # Remove so we can create it

    try:
        _, edit_result = await mcp.call_tool(
            "edit",
            {
                "file_path": str(new_path),
                "operation": "create",
            },
        )
        assert edit_result["success"]
        assert new_path.exists()

        # Verify it's readable
        _, read_result = await mcp.call_tool(
            "read", {"file_path": str(new_path), "scope": "blocks"}
        )
        assert read_result["block_count"] == 0
    finally:
        new_path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_edit_create_with_content():
    """Test creating a new document with initial content."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        new_path = Path(f.name)
    new_path.unlink()  # Remove so we can create it

    try:
        _, edit_result = await mcp.call_tool(
            "edit",
            {
                "file_path": str(new_path),
                "operation": "create",
                "content_type": "heading",
                "content_data": "My New Document",
                "heading_level": 1,
            },
        )
        assert edit_result["success"]
        assert edit_result["element_id"]
        assert new_path.exists()

        # Verify content
        _, read_result = await mcp.call_tool(
            "read", {"file_path": str(new_path), "scope": "blocks"}
        )
        assert read_result["block_count"] == 1
        assert read_result["blocks"][0]["type"] == "heading1"  # Level in type
        assert "My New Document" in read_result["blocks"][0]["text"]
    finally:
        new_path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_edit_create_fails_if_exists(sample_docx):
    """Test that create fails if file already exists."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "create",
        },
    )
    assert not edit_result["success"]
    assert "already exists" in edit_result["message"]


@pytest.mark.asyncio
async def test_block_ids_are_content_hash(sample_docx):
    """Test that block IDs use content-hash format (type_hash_occurrence)."""
    import re

    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    # Block IDs: paragraph_abc123_0, heading1_def456_0, table_ghi789_0
    id_pattern = re.compile(r"^(paragraph|heading[1-9]|table)_[0-9a-f]{8}_\d+$")
    for block in result["blocks"]:
        block_id = block["id"]
        assert id_pattern.match(block_id), f"Invalid ID format: {block_id}"


@pytest.mark.asyncio
async def test_read_table_cells(sample_docx):
    """Test reading table cells."""
    # First find the table
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")
    table_id = table_block["id"]

    # Now read the cells
    _, cells_result = await mcp.call_tool(
        "read",
        {"file_path": str(sample_docx), "scope": "table_cells", "target_id": table_id},
    )
    assert cells_result["table_rows"] == 2
    assert cells_result["table_cols"] == 2
    assert len(cells_result["cells"]) == 4  # 2x2 table
    # Check first cell
    cell_1_1 = next(c for c in cells_result["cells"] if c["row"] == 1 and c["col"] == 1)
    assert cell_1_1["text"] == "Header 1"


@pytest.mark.asyncio
async def test_edit_cell(sample_docx):
    """Test editing a single table cell."""
    # Find the table
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")
    table_id = table_block["id"]

    # Edit cell (1,1)
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "edit_cell",
            "target_id": table_id,
            "row": 1,
            "col": 1,
            "content_data": "Updated Header",
        },
    )
    assert edit_result["success"]
    # Returns updated table block ID (content-hash changes after cell edit)
    assert edit_result["element_id"].startswith("table_")
    assert edit_result["element_id"] != table_id  # Hash changed

    # Verify the change - content hash changes after edit, so re-read blocks
    _, blocks_result2 = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    new_table_block = next(b for b in blocks_result2["blocks"] if b["type"] == "table")
    new_table_id = new_table_block["id"]

    _, cells_result = await mcp.call_tool(
        "read",
        {
            "file_path": str(sample_docx),
            "scope": "table_cells",
            "target_id": new_table_id,
        },
    )
    cell_1_1 = next(c for c in cells_result["cells"] if c["row"] == 1 and c["col"] == 1)
    assert cell_1_1["text"] == "Updated Header"


@pytest.mark.asyncio
async def test_edit_cell_out_of_range(sample_docx):
    """Test that editing out-of-range cell fails."""
    # Find the table
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")
    table_id = table_block["id"]

    # Try to edit cell (99,1) - out of range
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "edit_cell",
            "target_id": table_id,
            "row": 99,
            "col": 1,
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "out of range" in edit_result["message"]


@pytest.mark.asyncio
async def test_edit_cell_on_non_table(sample_docx):
    """Test that editing cell on non-table block fails."""
    # Find a paragraph
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Try to edit cell on paragraph
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "edit_cell",
            "target_id": para_block["id"],
            "row": 1,
            "col": 1,
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "table" in edit_result["message"]
