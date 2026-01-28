"""Tests for list creation operations (Issue #148).

Tests create_list which bootstraps numbering.xml and applies numPr to paragraphs,
enabling bullet and numbered lists from scratch.
"""

import json
import tempfile
from pathlib import Path

import pytest

from mcp_handley_lab.microsoft.word.constants import qn
from mcp_handley_lab.microsoft.word.ops.lists import (
    create_list,
    get_list_info,
)
from mcp_handley_lab.microsoft.word.package import WordPackage
from mcp_handley_lab.microsoft.word.tool import mcp


@pytest.fixture
async def doc_with_paragraph():
    """Create a document with a single paragraph."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = Path(f.name)

    path.unlink()  # remove so edit() auto-creates
    await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps([{"op": "append", "content_data": "First item"}]),
        },
    )
    yield path
    path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_create_bullet_list(doc_with_paragraph):
    """create_list creates numPr on paragraph with bullet abstractNum."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    num_id = create_list(pkg, p_el, "bullet", level=0)
    pkg.save(str(path))

    assert num_id >= 1

    # Verify via get_list_info
    pkg2 = WordPackage.open(str(path))
    body2 = pkg2.body
    p_el2 = body2.find(qn("w:p"))
    info = get_list_info(pkg2, p_el2)

    assert info is not None
    assert info["num_id"] == num_id
    assert info["level"] == 0
    assert info["format_type"] == "bullet"


@pytest.mark.asyncio
async def test_create_numbered_list(doc_with_paragraph):
    """create_list with numbered type uses decimal numFmt at level 0."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_list(pkg, p_el, "numbered", level=0)
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    body2 = pkg2.body
    p_el2 = body2.find(qn("w:p"))
    info = get_list_info(pkg2, p_el2)

    assert info is not None
    assert info["format_type"] == "decimal"
    assert info["level_text"] == "%1."


@pytest.mark.asyncio
async def test_create_list_no_numbering_xml():
    """create_list bootstraps numbering.xml when it doesn't exist."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = Path(f.name)

    try:
        path.unlink()  # remove so edit() auto-creates
        await mcp.call_tool(
            "edit",
            {
                "file_path": str(path),
                "ops": json.dumps([{"op": "append", "content_data": "Item"}]),
            },
        )

        # Verify numbering.xml doesn't exist yet
        pkg = WordPackage.open(str(path))
        assert pkg.numbering_xml is None

        # Create list via MCP tool
        _, read_data = await mcp.call_tool(
            "read", {"file_path": str(path), "scope": "blocks"}
        )
        p_id = read_data["blocks"][0]["id"]

        _, edit_data = await mcp.call_tool(
            "edit",
            {
                "file_path": str(path),
                "ops": json.dumps(
                    [
                        {
                            "op": "create_list",
                            "target_id": p_id,
                            "content_data": json.dumps({"list_type": "bullet"}),
                        }
                    ]
                ),
            },
        )
        assert edit_data["results"][0]["success"] is True

        # Verify numbering.xml now exists
        pkg2 = WordPackage.open(str(path))
        assert pkg2.numbering_xml is not None
        assert pkg2.has_part("/word/numbering.xml")
    finally:
        path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_create_list_then_add_to_list(doc_with_paragraph):
    """Second item added via add_to_list uses the same numId."""
    path = doc_with_paragraph

    # Read paragraph ID
    _, read_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    p_id = read_data["blocks"][0]["id"]

    # Create list
    _, edit_data = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "create_list",
                        "target_id": p_id,
                        "content_data": json.dumps({"list_type": "bullet"}),
                    }
                ]
            ),
        },
    )
    new_p_id = edit_data["results"][0]["element_id"]

    # Add second item
    _, edit_data2 = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "add_to_list",
                        "target_id": new_p_id,
                        "content_data": json.dumps({"text": "Second item"}),
                    }
                ]
            ),
        },
    )
    assert edit_data2["results"][0]["success"] is True

    # Both paragraphs should have the same numId
    pkg = WordPackage.open(str(path))
    body = pkg.body
    paragraphs = body.findall(qn("w:p"))
    info1 = get_list_info(pkg, paragraphs[0])
    info2 = get_list_info(pkg, paragraphs[1])

    assert info1 is not None
    assert info2 is not None
    assert info1["num_id"] == info2["num_id"]


@pytest.mark.asyncio
async def test_create_list_roundtrip(doc_with_paragraph):
    """get_list_info reads back correctly after create_list."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    num_id = create_list(pkg, p_el, "numbered", level=2)
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    body2 = pkg2.body
    p_el2 = body2.find(qn("w:p"))
    info = get_list_info(pkg2, p_el2)

    assert info is not None
    assert info["num_id"] == num_id
    assert info["level"] == 2
    assert info["format_type"] == "lowerRoman"  # Level 2 cycles to lowerRoman
    assert info["level_text"] == "%3."


@pytest.mark.asyncio
async def test_create_list_level_validation(doc_with_paragraph):
    """Rejects level outside 0-8."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))

    with pytest.raises(ValueError, match="List level must be 0-8"):
        create_list(pkg, p_el, "bullet", level=9)

    with pytest.raises(ValueError, match="List level must be 0-8"):
        create_list(pkg, p_el, "bullet", level=-1)
