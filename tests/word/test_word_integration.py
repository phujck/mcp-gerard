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
    assert result["version"]  # Has version hash
    assert len(result["blocks"]) == 2  # Two headings
    assert result["blocks"][0]["type"] == "heading"
    assert result["blocks"][0]["level"] == 1
    assert "Test Document" in result["blocks"][0]["text"]
    assert result["blocks"][1]["level"] == 2


@pytest.mark.asyncio
async def test_read_meta(sample_docx):
    """Test reading document metadata."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "meta"}
    )
    assert result["version"]
    assert result["meta"] is not None
    assert result["blocks"] == []


@pytest.mark.asyncio
async def test_read_blocks(sample_docx):
    """Test reading all blocks."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    assert result["version"]
    assert result["block_count"] == 5  # 2 headings + 2 paragraphs + 1 table
    types = [b["type"] for b in result["blocks"]]
    assert "heading" in types
    assert "paragraph" in types
    assert "table" in types


@pytest.mark.asyncio
async def test_read_search(sample_docx):
    """Test searching for text."""
    _, result = await mcp.call_tool(
        "read",
        {"file_path": str(sample_docx), "scope": "search", "search_query": "second"},
    )
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
            "expected_version": read_result["version"],
            "operation": "append",
            "content_type": "paragraph",
            "content_data": "New paragraph at the end",
        },
    )
    assert edit_result["success"]
    assert edit_result["new_version"] != read_result["version"]

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
            "expected_version": read_result["version"],
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
async def test_edit_version_mismatch(sample_docx):
    """Test that version mismatch is detected."""
    _, result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "expected_version": "wrong_version",
            "operation": "append",
            "content_type": "paragraph",
            "content_data": "This should fail",
        },
    )
    assert not result["success"]
    assert (
        "modified" in result["message"].lower()
        or "re-read" in result["message"].lower()
    )


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
            "expected_version": read_result["version"],
            "operation": "insert_before",
            "target_id": target_id,
            "content_type": "paragraph",
            "content_data": "Inserted before",
        },
    )
    assert edit_result["success"]


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
            "expected_version": read_result["version"],
            "operation": "replace",
            "target_id": para_block["id"],
            "content_data": "Replaced content",
        },
    )
    assert edit_result["success"]


@pytest.mark.asyncio
async def test_edit_append_table(sample_docx):
    """Test appending a table."""
    _, read_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )

    table_data = json.dumps([["A", "B"], ["1", "2"]])
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "expected_version": read_result["version"],
            "operation": "append",
            "content_type": "table",
            "content_data": table_data,
        },
    )
    assert edit_result["success"]


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
        assert edit_result["new_version"]
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
        assert read_result["blocks"][0]["type"] == "heading"
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
