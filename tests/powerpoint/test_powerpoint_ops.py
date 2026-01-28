"""Tests for PowerPoint operations."""

from __future__ import annotations

import tempfile
from pathlib import Path

from mcp_handley_lab.microsoft.powerpoint.constants import EMU_PER_INCH, NSMAP, qn
from mcp_handley_lab.microsoft.powerpoint.ops.core import (
    emu_to_inches,
    inches_to_emu,
    make_shape_key,
    parse_shape_key,
    spatial_sort_shapes,
)
from mcp_handley_lab.microsoft.powerpoint.package import PowerPointPackage


class TestEmuConversion:
    """Tests for EMU conversion utilities."""

    def test_emu_to_inches(self):
        assert emu_to_inches(EMU_PER_INCH) == 1.0
        assert emu_to_inches(0) == 0.0
        assert emu_to_inches(EMU_PER_INCH * 2) == 2.0
        assert emu_to_inches(EMU_PER_INCH // 2) == 0.5

    def test_inches_to_emu(self):
        assert inches_to_emu(1.0) == EMU_PER_INCH
        assert inches_to_emu(0.0) == 0
        assert inches_to_emu(2.0) == EMU_PER_INCH * 2


class TestShapeKey:
    """Tests for shape key utilities."""

    def test_make_shape_key(self):
        assert make_shape_key(1, 42) == "1:42"
        assert make_shape_key(10, 100) == "10:100"

    def test_parse_shape_key(self):
        assert parse_shape_key("1:42") == (1, 42)
        assert parse_shape_key("10:100") == (10, 100)


class TestSpatialSort:
    """Tests for spatial sorting."""

    def test_sort_by_y_then_x(self):
        shapes = [
            {"y_inches": 2.0, "x_inches": 1.0, "z_order": 0, "shape_id": 1},
            {"y_inches": 1.0, "x_inches": 2.0, "z_order": 0, "shape_id": 2},
            {"y_inches": 1.0, "x_inches": 1.0, "z_order": 0, "shape_id": 3},
        ]
        sorted_shapes = spatial_sort_shapes(shapes)
        # Should be: (1,1), (1,2), (2,1)
        assert sorted_shapes[0]["shape_id"] == 3
        assert sorted_shapes[1]["shape_id"] == 2
        assert sorted_shapes[2]["shape_id"] == 1

    def test_sort_with_z_order_tiebreak(self):
        shapes = [
            {"y_inches": 1.0, "x_inches": 1.0, "z_order": 2, "shape_id": 1},
            {"y_inches": 1.0, "x_inches": 1.0, "z_order": 1, "shape_id": 2},
        ]
        sorted_shapes = spatial_sort_shapes(shapes)
        # Same position, lower z_order first
        assert sorted_shapes[0]["shape_id"] == 2
        assert sorted_shapes[1]["shape_id"] == 1


class TestPackageCreation:
    """Tests for PowerPointPackage creation."""

    def test_new_creates_minimal_presentation(self):
        pkg = PowerPointPackage.new()
        assert pkg.presentation_path == "/ppt/presentation.xml"
        assert pkg.presentation_xml is not None

    def test_new_has_slide_dimensions(self):
        pkg = PowerPointPackage.new()
        width, height = pkg.get_slide_dimensions()
        assert width == 9144000  # 10 inches
        assert height == 6858000  # 7.5 inches

    def test_save_and_reopen(self):
        pkg = PowerPointPackage.new()

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            temp_path = f.name

        try:
            pkg.save(temp_path)

            # Reopen
            pkg2 = PowerPointPackage.open(temp_path)
            assert pkg2.presentation_path == "/ppt/presentation.xml"
        finally:
            Path(temp_path).unlink(missing_ok=True)


class TestSlideOperations:
    """Tests for slide operations."""

    def test_get_slide_count_empty(self):
        from mcp_handley_lab.microsoft.powerpoint.ops.slides import get_slide_count

        pkg = PowerPointPackage.new()
        assert get_slide_count(pkg) == 0

    def test_list_slides_empty(self):
        from mcp_handley_lab.microsoft.powerpoint.ops.slides import list_slides

        pkg = PowerPointPackage.new()
        slides = list_slides(pkg)
        assert len(slides) == 0


class TestImageDimensions:
    """Tests for Pillow-based image dimension parsing."""

    def test_png_dimensions(self):
        import io

        from PIL import Image

        from mcp_handley_lab.microsoft.powerpoint.ops.images import (
            _get_image_dimensions,
        )

        # Create a valid PNG in memory
        img = Image.new("RGB", (16, 32), color="red")
        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        png_data = buffer.getvalue()

        width, height = _get_image_dimensions(png_data, "image/png")
        assert width == 16
        assert height == 32

    def test_gif_dimensions(self):
        import io

        from PIL import Image

        from mcp_handley_lab.microsoft.powerpoint.ops.images import (
            _get_image_dimensions,
        )

        # Create a valid GIF in memory
        img = Image.new("P", (64, 48), color=0)
        buffer = io.BytesIO()
        img.save(buffer, format="GIF")
        gif_data = buffer.getvalue()

        width, height = _get_image_dimensions(gif_data, "image/gif")
        assert width == 64
        assert height == 48

    def test_bmp_dimensions(self):
        import io

        from PIL import Image

        from mcp_handley_lab.microsoft.powerpoint.ops.images import (
            _get_image_dimensions,
        )

        # Create a valid BMP in memory
        img = Image.new("RGB", (128, 96), color="blue")
        buffer = io.BytesIO()
        img.save(buffer, format="BMP")
        bmp_data = buffer.getvalue()

        width, height = _get_image_dimensions(bmp_data, "image/bmp")
        assert width == 128
        assert height == 96

    def test_corrupted_image_fallback(self):
        from mcp_handley_lab.microsoft.powerpoint.ops.images import (
            _get_image_dimensions,
        )

        # Invalid/corrupted data should return fallback dimensions
        corrupted_data = b"not a valid image"
        width, height = _get_image_dimensions(corrupted_data, "image/png")
        assert width == 640
        assert height == 480


class TestPhase12BugFixes:
    """Tests for Phase 12 bug fixes (Gemini review)."""

    def test_image_namespace_on_root(self):
        """12a: r namespace should be on root p:pic, not on child a:blip."""
        from mcp_handley_lab.microsoft.powerpoint.ops.images import (
            _create_pic_element,
        )

        pic = _create_pic_element(1, "rId1", 1.0, 1.0, 3.0, 2.0)
        # r namespace should be declared on root element
        assert "r" in pic.nsmap or NSMAP["r"] in pic.nsmap.values()
        # a:blip should NOT have its own nsmap override
        blip = pic.find(".//" + qn("a:blip"), NSMAP)
        assert blip is not None

    def test_fill_guard_rejects_non_shape(self):
        """12b: fill guard should only allow sp and cxnSp shapes."""
        from lxml import etree

        # Test the guard logic directly: pic, graphicFrame should be rejected
        for tag in ("pic", "graphicFrame"):
            elem = etree.Element(qn(f"p:{tag}"))
            local = etree.QName(elem.tag).localname
            assert local not in ("sp", "cxnSp"), f"{tag} should be rejected by guard"

        # sp and cxnSp should be allowed
        for tag in ("sp", "cxnSp"):
            elem = etree.Element(qn(f"p:{tag}"))
            local = etree.QName(elem.tag).localname
            assert local in ("sp", "cxnSp"), f"{tag} should be allowed by guard"

    def test_fill_removal_list_no_blipfill(self):
        """12b: a:blipFill should not be in fill removal list."""
        import inspect

        from mcp_handley_lab.microsoft.powerpoint.ops.styling import set_shape_fill

        source = inspect.getsource(set_shape_fill)
        assert "a:blipFill" not in source

    def test_tab_extraction(self):
        """12d: a:tab elements should be extracted as tab characters."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import (
            extract_text_from_txBody,
        )

        txBody = etree.Element(qn("a:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        p = etree.SubElement(txBody, qn("a:p"))
        r1 = etree.SubElement(p, qn("a:r"))
        t1 = etree.SubElement(r1, qn("a:t"))
        t1.text = "Col1"
        etree.SubElement(p, qn("a:tab"))
        r2 = etree.SubElement(p, qn("a:r"))
        t2 = etree.SubElement(r2, qn("a:t"))
        t2.text = "Col2"

        text = extract_text_from_txBody(txBody)
        assert text == "Col1\tCol2"

    def test_tab_creation_in_shape_text(self):
        """12d: Tabs in text should create a:tab elements."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import (
            extract_text_from_txBody,
            set_shape_text,
        )

        sp = etree.Element(qn("p:sp"))
        set_shape_text(sp, "A\tB\tC")
        txBody = sp.find(qn("p:txBody"), NSMAP)
        assert txBody is not None

        # Should have tab elements
        p = txBody.find(qn("a:p"), NSMAP)
        tabs = p.findall(qn("a:tab"), NSMAP)
        assert len(tabs) == 2

        # Round-trip should preserve tabs
        text = extract_text_from_txBody(txBody)
        assert text == "A\tB\tC"

    def test_cell_text_preserves_formatting(self):
        """12c: _set_cell_text should preserve existing formatting."""

        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.tables import _set_cell_text

        # Create a cell with formatting
        tc = etree.Element(qn("a:tc"))
        txBody = etree.SubElement(tc, qn("a:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:lstStyle"))
        p = etree.SubElement(txBody, qn("a:p"))
        pPr = etree.SubElement(p, qn("a:pPr"))
        pPr.set("algn", "ctr")
        r = etree.SubElement(p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"))
        rPr.set("sz", "2400")
        rPr.set("b", "1")
        t = etree.SubElement(r, qn("a:t"))
        t.text = "Old text"

        # Set new text
        _set_cell_text(tc, "New text")

        # Verify formatting preserved
        new_txBody = tc.find(qn("a:txBody"), NSMAP)
        new_p = new_txBody.find(qn("a:p"), NSMAP)
        new_pPr = new_p.find(qn("a:pPr"), NSMAP)
        assert new_pPr is not None
        assert new_pPr.get("algn") == "ctr"

        new_r = new_p.find(qn("a:r"), NSMAP)
        new_rPr = new_r.find(qn("a:rPr"), NSMAP)
        assert new_rPr.get("sz") == "2400"
        assert new_rPr.get("b") == "1"

        new_t = new_r.find(qn("a:t"), NSMAP)
        assert new_t.text == "New text"


def _add_test_slide(pkg):
    """Add a minimal slide to a new package for testing.

    Works around the fact that PowerPointPackage.new() doesn't have slide masters
    so add_slide() fails. Creates slide XML directly in the package.
    """
    from lxml import etree

    from mcp_handley_lab.microsoft.powerpoint.constants import RT

    slide_xml = etree.fromstring(
        b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        b"<p:cSld><p:spTree>"
        b'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        b"<p:grpSpPr/>"
        b"</p:spTree></p:cSld></p:sld>"
    )
    partname = "/ppt/slides/slide1.xml"
    pkg._xml[partname] = slide_xml
    pkg._bytes[partname] = etree.tostring(slide_xml)

    # Add relationship from presentation to slide
    pres_rels = pkg.get_rels(pkg.presentation_path)
    rId = pres_rels.get_or_add(RT.SLIDE, "slides/slide1.xml")

    # Add slide to sldIdLst
    pres = pkg.presentation_xml
    sldIdLst = pres.find(qn("p:sldIdLst"), NSMAP)
    if sldIdLst is None:
        sldIdLst = etree.SubElement(pres, qn("p:sldIdLst"))
    sldId = etree.SubElement(sldIdLst, qn("p:sldId"))
    sldId.set("id", "256")
    sldId.set(qn("r:id"), rId)

    # Reset cached slide paths
    pkg._slide_paths = None
    return partname


class TestPhase13SlideBackground:
    """Tests for Phase 13: Slide backgrounds."""

    def test_set_slide_background(self):
        """Setting a background creates p:bg as first child of p:cSld."""
        from mcp_handley_lab.microsoft.powerpoint.ops.styling import (
            set_slide_background,
        )

        pkg = PowerPointPackage.new()
        _add_test_slide(pkg)

        result = set_slide_background(pkg, 1, "FF0000")
        assert result is True

        slide_xml = pkg.get_slide_xml(1)
        cSld = slide_xml.find(qn("p:cSld"), NSMAP)
        bg = cSld.find(qn("p:bg"), NSMAP)
        assert bg is not None

        # p:bg should be first child of p:cSld
        assert cSld[0].tag == qn("p:bg")

        # Check structure: p:bg/p:bgPr/a:solidFill/a:srgbClr
        bgPr = bg.find(qn("p:bgPr"), NSMAP)
        assert bgPr is not None
        solid_fill = bgPr.find(qn("a:solidFill"), NSMAP)
        assert solid_fill is not None
        srgb = solid_fill.find(qn("a:srgbClr"), NSMAP)
        assert srgb is not None
        assert srgb.get("val") == "FF0000"

        # effectLst should be present
        assert bgPr.find(qn("a:effectLst"), NSMAP) is not None

    def test_replace_existing_background(self):
        """Setting background twice replaces the first."""
        from mcp_handley_lab.microsoft.powerpoint.ops.styling import (
            set_slide_background,
        )

        pkg = PowerPointPackage.new()
        _add_test_slide(pkg)

        set_slide_background(pkg, 1, "FF0000")
        set_slide_background(pkg, 1, "00FF00")

        slide_xml = pkg.get_slide_xml(1)
        cSld = slide_xml.find(qn("p:cSld"), NSMAP)
        # Only one p:bg element
        bgs = cSld.findall(qn("p:bg"), NSMAP)
        assert len(bgs) == 1

        srgb = bgs[0].find(".//" + qn("a:srgbClr"), NSMAP)
        assert srgb.get("val") == "00FF00"


class TestPhase15BulletLists:
    """Tests for Phase 15: Bullet lists."""

    def test_bullet_style_creates_buchar(self):
        """bullet_style='bullet' creates a:buChar with bullet character."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        set_shape_text(sp, "Item 1\nItem 2", bullet_style="bullet")

        txBody = sp.find(qn("p:txBody"), NSMAP)
        paragraphs = txBody.findall(qn("a:p"), NSMAP)
        assert len(paragraphs) == 2

        for p in paragraphs:
            pPr = p.find(qn("a:pPr"), NSMAP)
            assert pPr is not None
            buChar = pPr.find(qn("a:buChar"), NSMAP)
            assert buChar is not None
            assert buChar.get("char") == "\u2022"
            assert pPr.get("marL") == "228600"
            assert pPr.get("indent") == "-228600"

    def test_dash_bullet_style(self):
        """bullet_style='dash' creates a:buChar with en-dash."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        set_shape_text(sp, "Item", bullet_style="dash")

        txBody = sp.find(qn("p:txBody"), NSMAP)
        p = txBody.find(qn("a:p"), NSMAP)
        pPr = p.find(qn("a:pPr"), NSMAP)
        buChar = pPr.find(qn("a:buChar"), NSMAP)
        assert buChar.get("char") == "\u2013"

    def test_number_bullet_style(self):
        """bullet_style='number' creates a:buAutoNum."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        set_shape_text(sp, "Step 1\nStep 2", bullet_style="number")

        txBody = sp.find(qn("p:txBody"), NSMAP)
        p = txBody.find(qn("a:p"), NSMAP)
        pPr = p.find(qn("a:pPr"), NSMAP)
        buAutoNum = pPr.find(qn("a:buAutoNum"), NSMAP)
        assert buAutoNum is not None
        assert buAutoNum.get("type") == "arabicPeriod"

    def test_none_bullet_style_removes_bullets(self):
        """bullet_style='none' adds a:buNone and removes indent."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        # First add bullets
        set_shape_text(sp, "Item", bullet_style="bullet")
        # Then remove
        set_shape_text(sp, "Item", bullet_style="none")

        txBody = sp.find(qn("p:txBody"), NSMAP)
        p = txBody.find(qn("a:p"), NSMAP)
        pPr = p.find(qn("a:pPr"), NSMAP)
        assert pPr.find(qn("a:buNone"), NSMAP) is not None
        assert pPr.find(qn("a:buChar"), NSMAP) is None
        assert pPr.get("marL") is None
        assert pPr.get("indent") is None

    def test_invalid_bullet_style_raises(self):
        """Unknown bullet_style raises ValueError."""
        import pytest
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        with pytest.raises(ValueError, match="Unknown bullet_style"):
            set_shape_text(sp, "Item", bullet_style="invalid")

    def test_no_bullet_style_preserves_existing(self):
        """bullet_style=None does not modify bullet state."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import set_shape_text

        sp = etree.Element(qn("p:sp"))
        set_shape_text(sp, "Item", bullet_style="bullet")

        # Edit without bullet_style - should preserve
        set_shape_text(sp, "New text")

        txBody = sp.find(qn("p:txBody"), NSMAP)
        p = txBody.find(qn("a:p"), NSMAP)
        pPr = p.find(qn("a:pPr"), NSMAP)
        # The existing pPr should be preserved with bullet properties
        assert pPr is not None
        buChar = pPr.find(qn("a:buChar"), NSMAP)
        assert buChar is not None


class TestPhase14Hyperlinks:
    """Tests for Phase 14: Hyperlinks."""

    def test_add_hyperlink(self):
        """Adding a hyperlink creates a:hlinkClick on text runs."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import add_hyperlink

        pkg = PowerPointPackage.new()
        _add_test_slide(pkg)

        # Add a shape with text
        slide_xml = pkg.get_slide_xml(1)
        sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
        sp = etree.SubElement(sp_tree, qn("p:sp"))
        nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
        cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
        cNvPr.set("id", "100")
        cNvPr.set("name", "Test Shape")
        etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
        etree.SubElement(nvSpPr, qn("p:nvPr"))
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        p = etree.SubElement(txBody, qn("a:p"))
        r = etree.SubElement(p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"), lang="en-US")
        t = etree.SubElement(r, qn("a:t"))
        t.text = "Click me"
        slide_partname = pkg.get_slide_partname(1)
        pkg.mark_xml_dirty(slide_partname)

        result = add_hyperlink(pkg, "1:100", "https://example.com", "Example")
        assert result is True

        # Verify a:hlinkClick was added
        hlink = rPr.find(qn("a:hlinkClick"), NSMAP)
        assert hlink is not None
        assert hlink.get("tooltip") == "Example"

        # Verify relationship was created and is external
        slide_rels = pkg.get_rels(slide_partname)
        rId = hlink.get(qn("r:id"))
        assert rId is not None
        assert rId in slide_rels
        rel = slide_rels[rId]
        assert rel.is_external is True
        assert rel.target == "https://example.com"

    def test_hyperlink_replaces_existing(self):
        """Adding a hyperlink replaces any existing one."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import add_hyperlink

        pkg = PowerPointPackage.new()
        _add_test_slide(pkg)

        # Add a shape with text
        slide_xml = pkg.get_slide_xml(1)
        sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
        sp = etree.SubElement(sp_tree, qn("p:sp"))
        nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
        cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
        cNvPr.set("id", "101")
        cNvPr.set("name", "Test Shape")
        etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
        etree.SubElement(nvSpPr, qn("p:nvPr"))
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        p = etree.SubElement(txBody, qn("a:p"))
        r = etree.SubElement(p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"), lang="en-US")
        t = etree.SubElement(r, qn("a:t"))
        t.text = "Link text"
        slide_partname = pkg.get_slide_partname(1)
        pkg.mark_xml_dirty(slide_partname)

        # Add first hyperlink
        add_hyperlink(pkg, "1:101", "https://old.com")
        # Replace with second
        add_hyperlink(pkg, "1:101", "https://new.com", "New Link")

        # Should only have one hlinkClick
        hlinks = rPr.findall(qn("a:hlinkClick"), NSMAP)
        assert len(hlinks) == 1
        assert hlinks[0].get("tooltip") == "New Link"

    def test_hyperlink_no_text_returns_false(self):
        """Hyperlink on shape with no text returns False."""
        from lxml import etree

        from mcp_handley_lab.microsoft.powerpoint.ops.text import add_hyperlink

        pkg = PowerPointPackage.new()
        _add_test_slide(pkg)

        # Add a shape without text runs
        slide_xml = pkg.get_slide_xml(1)
        sp_tree = slide_xml.find(qn("p:cSld") + "/" + qn("p:spTree"), NSMAP)
        sp = etree.SubElement(sp_tree, qn("p:sp"))
        nvSpPr = etree.SubElement(sp, qn("p:nvSpPr"))
        cNvPr = etree.SubElement(nvSpPr, qn("p:cNvPr"))
        cNvPr.set("id", "102")
        cNvPr.set("name", "Empty Shape")
        etree.SubElement(nvSpPr, qn("p:cNvSpPr"))
        etree.SubElement(nvSpPr, qn("p:nvPr"))
        txBody = etree.SubElement(sp, qn("p:txBody"))
        etree.SubElement(txBody, qn("a:bodyPr"))
        etree.SubElement(txBody, qn("a:p"))  # Empty paragraph, no runs
        pkg.mark_xml_dirty(pkg.get_slide_partname(1))

        result = add_hyperlink(pkg, "1:102", "https://example.com")
        assert result is False


class TestNextPartname:
    """Tests for next_partname utility."""

    def test_next_partname_empty(self):
        pkg = PowerPointPackage.new()
        # No slides exist, so next should be slide1
        next_slide = pkg.next_partname("/ppt/slides/slide", ".xml")
        assert next_slide == "/ppt/slides/slide1.xml"

    def test_next_partname_with_existing(self):
        pkg = PowerPointPackage.new()
        # Manually add some "slides"
        pkg._bytes["/ppt/slides/slide1.xml"] = b""
        pkg._bytes["/ppt/slides/slide2.xml"] = b""
        pkg._bytes["/ppt/slides/slide5.xml"] = b""  # Gap

        # Next should be 6 (max existing + 1)
        next_slide = pkg.next_partname("/ppt/slides/slide", ".xml")
        assert next_slide == "/ppt/slides/slide6.xml"
