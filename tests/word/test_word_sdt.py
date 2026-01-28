"""Tests for content control (SDT) creation (Issue #126).

Tests create_content_control for text, checkbox, dropdown, comboBox, and date types.
"""

import json
import tempfile
from pathlib import Path

import pytest
from lxml import etree

from mcp_handley_lab.microsoft.word.constants import NSMAP, qn
from mcp_handley_lab.microsoft.word.ops.sdt import (
    build_content_controls,
    create_content_control,
)
from mcp_handley_lab.microsoft.word.package import WordPackage
from mcp_handley_lab.microsoft.word.tool import mcp


@pytest.fixture
async def doc_with_paragraph():
    """Create a document with a single paragraph."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = Path(f.name)

    await mcp.call_tool(
        "create",
        {
            "file_path": str(path),
            "content_type": "paragraph",
            "content_data": "Reference paragraph",
        },
    )
    yield path
    path.unlink(missing_ok=True)


@pytest.mark.asyncio
async def test_create_text_control(doc_with_paragraph):
    """Create a text content control and read it back."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    sdt = create_content_control(pkg, body, p_el, "text", tag="myTag", alias="My Alias")
    pkg.save(str(path))

    assert sdt is not None

    # Read back via build_content_controls
    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert len(controls) == 1
    cc = controls[0]
    assert cc["type"] == "text"
    assert cc["tag"] == "myTag"
    assert cc["alias"] == "My Alias"
    assert cc["value"] == "Click here"  # Default placeholder


@pytest.mark.asyncio
async def test_create_checkbox(doc_with_paragraph):
    """Create a checkbox control, verify checked state and structural elements."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_content_control(pkg, body, p_el, "checkbox", checked=True)
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert len(controls) == 1
    cc = controls[0]
    assert cc["type"] == "checkbox"
    assert cc["checked"] is True
    assert cc["value"] == "\u2612"  # ☒

    # Verify structural elements in XML
    body2 = pkg2.body
    sdt_el = list(body2.iter(qn("w:sdt")))[0]
    sdt_pr = sdt_el.find("w:sdtPr", namespaces=NSMAP)
    ns_w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    checkbox = sdt_pr.find(f"{{{ns_w14}}}checkbox")
    assert checkbox is not None

    checked_state = checkbox.find(f"{{{ns_w14}}}checkedState")
    assert checked_state is not None
    assert checked_state.get(f"{{{ns_w14}}}val") == "2612"

    unchecked_state = checkbox.find(f"{{{ns_w14}}}uncheckedState")
    assert unchecked_state is not None
    assert unchecked_state.get(f"{{{ns_w14}}}val") == "2610"


@pytest.mark.asyncio
async def test_create_checkbox_unchecked(doc_with_paragraph):
    """Create an unchecked checkbox."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_content_control(pkg, body, p_el, "checkbox", checked=False)
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert controls[0]["checked"] is False
    assert controls[0]["value"] == "\u2610"  # ☐


@pytest.mark.asyncio
async def test_create_dropdown(doc_with_paragraph):
    """Create a dropdown control with options."""
    path = doc_with_paragraph
    options = ["Red", "Green", "Blue"]

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_content_control(pkg, body, p_el, "dropdown", options=options, tag="color")
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert len(controls) == 1
    cc = controls[0]
    assert cc["type"] == "dropdown"
    assert cc["options"] == ["Red", "Green", "Blue"]
    assert cc["tag"] == "color"


@pytest.mark.asyncio
async def test_create_combobox(doc_with_paragraph):
    """Create a comboBox control with options."""
    path = doc_with_paragraph
    options = ["Option A", "Option B"]

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_content_control(pkg, body, p_el, "comboBox", options=options)
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert len(controls) == 1
    cc = controls[0]
    assert cc["type"] == "comboBox"
    assert cc["options"] == ["Option A", "Option B"]


@pytest.mark.asyncio
async def test_create_date(doc_with_paragraph):
    """Create a date control with custom format."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    create_content_control(pkg, body, p_el, "date", date_format="dd/MM/yyyy")
    pkg.save(str(path))

    pkg2 = WordPackage.open(str(path))
    controls = build_content_controls(pkg2)

    assert len(controls) == 1
    cc = controls[0]
    assert cc["type"] == "date"
    assert cc["date_format"] == "dd/MM/yyyy"

    # Verify XML structure
    body2 = pkg2.body
    sdt_el = list(body2.iter(qn("w:sdt")))[0]
    sdt_pr = sdt_el.find("w:sdtPr", namespaces=NSMAP)
    date_el = sdt_pr.find("w:date", namespaces=NSMAP)
    assert date_el is not None

    lid = date_el.find("w:lid", namespaces=NSMAP)
    assert lid is not None
    assert lid.get(qn("w:val")) == "en-US"

    store = date_el.find("w:storeMappedDataAs", namespaces=NSMAP)
    assert store is not None
    assert store.get(qn("w:val")) == "dateTime"


@pytest.mark.asyncio
async def test_sdt_id_uniqueness(doc_with_paragraph):
    """Multiple creates produce unique IDs."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body
    p_el = body.find(qn("w:p"))
    sdt1 = create_content_control(pkg, body, p_el, "text", tag="first")
    sdt2 = create_content_control(pkg, body, p_el, "text", tag="second")
    pkg.save(str(path))

    # Extract IDs
    sdt1_pr = sdt1.find("w:sdtPr", namespaces=NSMAP)
    sdt2_pr = sdt2.find("w:sdtPr", namespaces=NSMAP)
    id1 = int(sdt1_pr.find("w:id", namespaces=NSMAP).get(qn("w:val")))
    id2 = int(sdt2_pr.find("w:id", namespaces=NSMAP).get(qn("w:val")))
    assert id1 != id2


@pytest.mark.asyncio
async def test_create_sdt_parent_validation(doc_with_paragraph):
    """Only block-level contexts (w:body, w:tc) are accepted."""
    path = doc_with_paragraph

    pkg = WordPackage.open(str(path))
    body = pkg.body

    # Create a detached paragraph (no parent)
    detached_p = etree.Element(qn("w:p"))

    with pytest.raises(ValueError, match="block-level"):
        create_content_control(pkg, body, detached_p, "text")


@pytest.mark.asyncio
async def test_create_content_control_via_mcp(doc_with_paragraph):
    """Test create_content_control via the MCP tool interface."""
    path = doc_with_paragraph

    # Read paragraph ID
    _, read_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "blocks"}
    )
    p_id = read_data["blocks"][0]["id"]

    # Create via MCP
    _, edit_data = await mcp.call_tool(
        "edit",
        {
            "file_path": str(path),
            "ops": json.dumps(
                [
                    {
                        "op": "create_content_control",
                        "target_id": p_id,
                        "content_data": json.dumps(
                            {
                                "type": "dropdown",
                                "tag": "priority",
                                "options": ["Low", "Medium", "High"],
                            }
                        ),
                    }
                ]
            ),
        },
    )

    assert edit_data["results"][0]["success"] is True
    # element_id is the numeric SDT id (usable with set_content_control)
    assert edit_data["results"][0]["element_id"].isdigit()

    # Verify via read
    _, cc_data = await mcp.call_tool(
        "read", {"file_path": str(path), "scope": "content_controls"}
    )
    assert len(cc_data["content_controls"]) == 1
    assert cc_data["content_controls"][0]["type"] == "dropdown"
    assert cc_data["content_controls"][0]["options"] == ["Low", "Medium", "High"]
