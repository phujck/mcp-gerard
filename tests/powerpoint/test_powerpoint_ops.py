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
