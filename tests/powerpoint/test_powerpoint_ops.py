"""Tests for PowerPoint operations."""

from __future__ import annotations

import tempfile
from pathlib import Path

from mcp_handley_lab.microsoft.powerpoint.constants import EMU_PER_INCH
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
