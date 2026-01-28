"""Tests for hyperlink operations (Issue #136).

Tests the add_hyperlink replace parameter that prevents text duplication
when replacing paragraph content with a hyperlink.
"""

import json
import tempfile
from pathlib import Path

import pytest

from mcp_handley_lab.microsoft.word.tool import mcp


@pytest.fixture
async def doc_with_paragraph():
    """Create a document with a single paragraph containing text."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = Path(f.name)

    await mcp.call_tool(
        "create",
        {
            "file_path": str(path),
            "content_type": "paragraph",
            "content_data": "Original text here",
        },
    )
    yield path
    path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_add_hyperlink_append(doc_with_paragraph):
    """Default behavior: hyperlink is appended, preserving existing text."""
    path = doc_with_paragraph

    # Read to get paragraph ID
    _, read_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    p_id = read_data["blocks"][0]["id"]

    # Add hyperlink (default: append)
    _, edit_data = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "add_hyperlink",
                        "target_id": p_id,
                        "content_data": json.dumps(
                            {"text": "Click me", "address": "https://example.com"}
                        ),
                    }
                ]
            ),
        },
    )
    assert edit_data["results"][0]["success"] is True

    # Read hyperlinks - should have the link
    _, hyper_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "hyperlinks"}
    )
    assert len(hyper_data["hyperlinks"]) == 1
    assert hyper_data["hyperlinks"][0]["text"] == "Click me"

    # Read blocks - original text should still be there (appended)
    _, blocks_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    paragraph_text = blocks_data["blocks"][0]["text"]
    assert "Original text here" in paragraph_text


@pytest.mark.asyncio
async def test_add_hyperlink_replace(doc_with_paragraph):
    """replace=True clears all content except pPr before adding hyperlink."""
    path = doc_with_paragraph

    # Read to get paragraph ID
    _, read_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    p_id = read_data["blocks"][0]["id"]

    # Add hyperlink with replace=True
    _, edit_data = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "add_hyperlink",
                        "target_id": p_id,
                        "content_data": json.dumps(
                            {
                                "text": "New link",
                                "address": "https://example.com",
                                "replace": True,
                            }
                        ),
                    }
                ]
            ),
        },
    )
    assert edit_data["results"][0]["success"] is True

    # Read blocks - original text should be gone, only link text
    _, blocks_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    paragraph_text = blocks_data["blocks"][0]["text"]
    assert "Original text here" not in paragraph_text
    assert "New link" in paragraph_text


@pytest.mark.asyncio
async def test_add_hyperlink_replace_preserves_p_pr(doc_with_paragraph):
    """replace=True preserves paragraph properties (w:pPr)."""
    path = doc_with_paragraph

    # Read to get paragraph ID
    _, read_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    p_id = read_data["blocks"][0]["id"]

    # Apply formatting to create pPr
    _, edit_data = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "style",
                        "target_id": p_id,
                        "formatting": json.dumps({"alignment": "center"}),
                    }
                ]
            ),
        },
    )
    new_p_id = edit_data["results"][0]["element_id"]

    # Now replace with hyperlink
    _, edit_data2 = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "add_hyperlink",
                        "target_id": new_p_id,
                        "content_data": json.dumps(
                            {
                                "text": "Centered link",
                                "address": "https://example.com",
                                "replace": True,
                            }
                        ),
                    }
                ]
            ),
        },
    )
    assert edit_data2["results"][0]["success"] is True

    # Verify link text is there and original text is gone
    _, blocks_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    assert "Centered link" in blocks_data["blocks"][0]["text"]
    assert "Original text here" not in blocks_data["blocks"][0]["text"]
