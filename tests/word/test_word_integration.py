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


@pytest.fixture
def formatted_docx():
    """Create a Word document with formatted runs for testing."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        doc = Document()
        # Create a paragraph with multiple runs with different formatting
        p = doc.add_paragraph()
        p.add_run("Normal text, ")  # default formatting
        run2 = p.add_run("bold text, ")
        run2.bold = True
        run3 = p.add_run("italic text.")
        run3.italic = True
        doc.save(f.name)
        yield Path(f.name)
    Path(f.name).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_read_runs(formatted_docx):
    """Test reading runs from a paragraph."""
    # First get the paragraph ID
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Now read the runs
    _, runs_result = await mcp.call_tool(
        "read",
        {
            "file_path": str(formatted_docx),
            "scope": "runs",
            "target_id": para_block["id"],
        },
    )
    assert runs_result["block_count"] == 3  # Three runs
    assert len(runs_result["runs"]) == 3

    # Check run properties
    assert runs_result["runs"][0]["text"] == "Normal text, "
    assert runs_result["runs"][0]["bold"] is None  # Not set (inherits)

    assert runs_result["runs"][1]["text"] == "bold text, "
    assert runs_result["runs"][1]["bold"] is True

    assert runs_result["runs"][2]["text"] == "italic text."
    assert runs_result["runs"][2]["italic"] is True


@pytest.mark.asyncio
async def test_edit_run_text(formatted_docx):
    """Test editing run text."""
    # Get paragraph ID
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Edit the bold run (index 1) - use different text (not just case change, as hash normalizes)
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(formatted_docx),
            "operation": "edit_run",
            "target_id": para_block["id"],
            "run_index": 1,
            "content_data": "strongly emphasized, ",
        },
    )
    assert edit_result["success"]
    # Hash changes because paragraph text changed
    assert edit_result["element_id"].startswith("paragraph_")
    assert edit_result["element_id"] != para_block["id"]

    # Verify the change - need new paragraph ID
    _, blocks_result2 = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    new_para = next(b for b in blocks_result2["blocks"] if b["type"] == "paragraph")

    _, runs_result = await mcp.call_tool(
        "read",
        {
            "file_path": str(formatted_docx),
            "scope": "runs",
            "target_id": new_para["id"],
        },
    )
    assert runs_result["runs"][1]["text"] == "strongly emphasized, "
    assert runs_result["runs"][1]["bold"] is True  # Formatting preserved


@pytest.mark.asyncio
async def test_edit_run_formatting(formatted_docx):
    """Test editing run formatting."""
    # Get paragraph ID
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Make the first run (Normal text) bold and underlined
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(formatted_docx),
            "operation": "edit_run",
            "target_id": para_block["id"],
            "run_index": 0,
            "formatting": '{"bold": true, "underline": true}',
        },
    )
    assert edit_result["success"]

    # Verify formatting (text unchanged, so ID unchanged)
    _, runs_result = await mcp.call_tool(
        "read",
        {
            "file_path": str(formatted_docx),
            "scope": "runs",
            "target_id": para_block["id"],
        },
    )
    assert runs_result["runs"][0]["text"] == "Normal text, "
    assert runs_result["runs"][0]["bold"] is True
    assert runs_result["runs"][0]["underline"] is True


@pytest.mark.asyncio
async def test_edit_run_out_of_range(formatted_docx):
    """Test that editing out-of-range run fails."""
    # Get paragraph ID
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Try to edit run 99 - out of range
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(formatted_docx),
            "operation": "edit_run",
            "target_id": para_block["id"],
            "run_index": 99,
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "out of range" in edit_result["message"]


@pytest.mark.asyncio
async def test_read_runs_on_table_fails(sample_docx):
    """Test that reading runs on a table fails."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")

    _, runs_result = await mcp.call_tool(
        "read",
        {
            "file_path": str(sample_docx),
            "scope": "runs",
            "target_id": table_block["id"],
        },
    )
    assert runs_result["block_count"] == 0
    assert "paragraph or heading" in runs_result["warnings"][0]


@pytest.mark.asyncio
async def test_edit_run_on_table_fails(sample_docx):
    """Test that editing run on a table fails."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "edit_run",
            "target_id": table_block["id"],
            "run_index": 0,
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "paragraph or heading" in edit_result["message"]


@pytest.mark.asyncio
async def test_edit_run_missing_run_index(formatted_docx):
    """Test that edit_run without run_index fails."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(formatted_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(formatted_docx),
            "operation": "edit_run",
            "target_id": para_block["id"],
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "run_index" in edit_result["message"]


# --- Comment tests ---


@pytest.mark.asyncio
async def test_read_comments_empty(sample_docx):
    """Test reading comments from a document with no comments."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "comments"}
    )
    assert result["block_count"] == 0
    assert result["comments"] == []


@pytest.mark.asyncio
async def test_add_comment_to_paragraph(sample_docx):
    """Test adding a comment to a paragraph."""
    # First get a paragraph ID
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Add a comment
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "add_comment",
            "target_id": para_block["id"],
            "content_data": "This is a test comment",
        },
    )
    assert edit_result["success"]
    assert edit_result["comment_id"] is not None

    # Verify comment was added
    _, comments_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "comments"}
    )
    assert comments_result["block_count"] == 1
    assert len(comments_result["comments"]) == 1
    assert comments_result["comments"][0]["text"] == "This is a test comment"


@pytest.mark.asyncio
async def test_add_comment_with_author(sample_docx):
    """Test adding a comment with author and initials."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    # Add a comment with author
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "add_comment",
            "target_id": para_block["id"],
            "content_data": "Review needed",
            "author": "Test Author",
            "initials": "TA",
        },
    )
    assert edit_result["success"]

    # Verify author info
    _, comments_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "comments"}
    )
    assert comments_result["comments"][0]["author"] == "Test Author"
    assert comments_result["comments"][0]["initials"] == "TA"


@pytest.mark.asyncio
async def test_add_comment_to_table_fails(sample_docx):
    """Test that adding a comment to a table fails."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    table_block = next(b for b in blocks_result["blocks"] if b["type"] == "table")

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "add_comment",
            "target_id": table_block["id"],
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "paragraph or heading" in edit_result["message"]


@pytest.mark.asyncio
async def test_add_comment_missing_content(sample_docx):
    """Test that add_comment without content fails."""
    _, blocks_result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "blocks"}
    )
    para_block = next(b for b in blocks_result["blocks"] if b["type"] == "paragraph")

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "add_comment",
            "target_id": para_block["id"],
        },
    )
    assert not edit_result["success"]
    assert "content_data" in edit_result["message"]


# --- Headers/Footers tests ---


@pytest.mark.asyncio
async def test_read_headers_footers(sample_docx):
    """Test reading headers and footers."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "headers_footers"}
    )
    # Document should have at least one section
    assert result["block_count"] >= 1
    assert len(result["headers_footers"]) >= 1
    # First section should be at index 0
    assert result["headers_footers"][0]["section_index"] == 0


@pytest.mark.asyncio
async def test_set_header(sample_docx):
    """Test setting a header."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_header",
            "section_index": 0,
            "content_data": "My Custom Header",
        },
    )
    assert edit_result["success"]

    # Verify header was set
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "headers_footers"}
    )
    assert result["headers_footers"][0]["header_text"] == "My Custom Header"
    assert result["headers_footers"][0]["header_is_linked"] is False


@pytest.mark.asyncio
async def test_set_footer(sample_docx):
    """Test setting a footer."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_footer",
            "section_index": 0,
            "content_data": "Page Footer Text",
        },
    )
    assert edit_result["success"]

    # Verify footer was set
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "headers_footers"}
    )
    assert result["headers_footers"][0]["footer_text"] == "Page Footer Text"
    assert result["headers_footers"][0]["footer_is_linked"] is False


@pytest.mark.asyncio
async def test_set_header_invalid_section(sample_docx):
    """Test that setting header on invalid section fails."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_header",
            "section_index": 99,
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "out of range" in edit_result["message"]


# --- Page Setup tests ---


@pytest.mark.asyncio
async def test_read_page_setup(sample_docx):
    """Test reading page setup."""
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "page_setup"}
    )
    # Document should have at least one section
    assert result["block_count"] >= 1
    assert len(result["page_setup"]) >= 1
    # Check first section has expected fields
    setup = result["page_setup"][0]
    assert setup["section_index"] == 0
    assert setup["orientation"] in ["portrait", "landscape"]
    assert setup["page_width"] > 0
    assert setup["page_height"] > 0


@pytest.mark.asyncio
async def test_set_margins(sample_docx):
    """Test setting page margins."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_margins",
            "section_index": 0,
            "formatting": '{"top": 0.5, "bottom": 0.5, "left": 0.75, "right": 0.75}',
        },
    )
    assert edit_result["success"]

    # Verify margins were set
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "page_setup"}
    )
    setup = result["page_setup"][0]
    assert setup["top_margin"] == 0.5
    assert setup["bottom_margin"] == 0.5
    assert setup["left_margin"] == 0.75
    assert setup["right_margin"] == 0.75


@pytest.mark.asyncio
async def test_set_orientation_landscape(sample_docx):
    """Test setting page orientation to landscape."""
    # Get original dimensions
    _, before = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "page_setup"}
    )
    orig_width = before["page_setup"][0]["page_width"]
    orig_height = before["page_setup"][0]["page_height"]

    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_orientation",
            "section_index": 0,
            "content_data": "landscape",
        },
    )
    assert edit_result["success"]

    # Verify orientation was set
    _, result = await mcp.call_tool(
        "read", {"file_path": str(sample_docx), "scope": "page_setup"}
    )
    setup = result["page_setup"][0]
    assert setup["orientation"] == "landscape"
    # Width should now be greater than height (or swapped)
    assert setup["page_width"] >= setup["page_height"] or (
        setup["page_width"] == orig_height and setup["page_height"] == orig_width
    )


@pytest.mark.asyncio
async def test_set_orientation_invalid(sample_docx):
    """Test that invalid orientation fails."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_orientation",
            "section_index": 0,
            "content_data": "diagonal",
        },
    )
    assert not edit_result["success"]
    assert "Invalid orientation" in edit_result["message"]


@pytest.mark.asyncio
async def test_set_margins_missing_formatting(sample_docx):
    """Test that set_margins without formatting fails."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": str(sample_docx),
            "operation": "set_margins",
            "section_index": 0,
        },
    )
    assert not edit_result["success"]
    assert "formatting" in edit_result["message"]


# --- File error handling tests ---


@pytest.mark.asyncio
async def test_read_missing_file():
    """Test that reading a missing file returns structured error."""
    _, result = await mcp.call_tool(
        "read",
        {
            "file_path": "/nonexistent/path/to/file.docx",
            "scope": "blocks",
        },
    )
    assert result["block_count"] == 0
    assert len(result["warnings"]) > 0
    assert "not found" in result["warnings"][0].lower()


@pytest.mark.asyncio
async def test_edit_missing_file():
    """Test that editing a missing file returns structured error."""
    _, edit_result = await mcp.call_tool(
        "edit",
        {
            "file_path": "/nonexistent/path/to/file.docx",
            "operation": "append",
            "content_data": "Should fail",
        },
    )
    assert not edit_result["success"]
    assert "not found" in edit_result["message"].lower()


@pytest.mark.asyncio
async def test_read_corrupt_file():
    """Test that reading a corrupt file returns structured error."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        f.write(b"not a valid docx file")
        corrupt_path = f.name
    try:
        _, result = await mcp.call_tool(
            "read",
            {
                "file_path": corrupt_path,
                "scope": "blocks",
            },
        )
        assert result["block_count"] == 0
        assert len(result["warnings"]) > 0
        # python-docx may raise PackageNotFoundError for corrupt files too
        warning_lower = result["warnings"][0].lower()
        assert any(msg in warning_lower for msg in ["corrupt", "invalid", "not found"])
    finally:
        Path(corrupt_path).unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_create_in_new_directory():
    """Test creating a document in a new (non-existent) directory."""
    with tempfile.TemporaryDirectory() as tmpdir:
        # Create path with nested non-existent directory
        new_file = Path(tmpdir) / "subdir" / "nested" / "test.docx"
        _, result = await mcp.call_tool(
            "edit",
            {
                "file_path": str(new_file),
                "operation": "create",
                "content_data": "Test content",
            },
        )
        assert result["success"]
        assert new_file.exists()
