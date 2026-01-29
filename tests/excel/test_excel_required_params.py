"""Tests for required parameter validation in Excel edit operations."""

import json

import pytest

from mcp_handley_lab.microsoft.excel.package import ExcelPackage
from mcp_handley_lab.microsoft.excel.shared import edit


@pytest.fixture
def workbook(tmp_path):
    """Create a temporary Excel workbook with a sheet."""
    file_path = tmp_path / "test.xlsx"
    pkg = ExcelPackage.new()
    pkg.save(str(file_path))
    return str(file_path)


class TestCellOpsRequiredParams:
    """Validate required params for cell operations."""

    def test_set_cell_missing_sheet(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps([{"op": "set_cell", "cell_ref": "A1", "value": "x"}]),
        )
        assert result["results"][0]["success"] is False
        assert "sheet required" in result["results"][0]["error"].lower()

    def test_set_cell_missing_cell_ref(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps([{"op": "set_cell", "sheet": "Sheet1", "value": "x"}]),
        )
        assert result["results"][0]["success"] is False
        assert "cell_ref required" in result["results"][0]["error"].lower()

    def test_set_cell_missing_value(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps([{"op": "set_cell", "sheet": "Sheet1", "cell_ref": "A1"}]),
        )
        assert result["results"][0]["success"] is False
        assert "value required" in result["results"][0]["error"].lower()

    def test_set_formula_missing_formula(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps(
                [{"op": "set_formula", "sheet": "Sheet1", "cell_ref": "A1"}]
            ),
        )
        assert result["results"][0]["success"] is False
        assert "formula required" in result["results"][0]["error"].lower()


class TestSheetOpsRequiredParams:
    """Validate required params for sheet operations."""

    def test_add_sheet_missing_name(self, workbook):
        result = edit(workbook, ops=json.dumps([{"op": "add_sheet"}]))
        assert result["results"][0]["success"] is False
        assert "required for add_sheet" in result["results"][0]["error"].lower()

    def test_rename_sheet_missing_new_name(self, workbook):
        result = edit(
            workbook, ops=json.dumps([{"op": "rename_sheet", "sheet": "Sheet1"}])
        )
        assert result["results"][0]["success"] is False
        assert "new_name required" in result["results"][0]["error"].lower()

    def test_delete_sheet_missing_sheet(self, workbook):
        result = edit(workbook, ops=json.dumps([{"op": "delete_sheet"}]))
        assert result["results"][0]["success"] is False
        assert "sheet required" in result["results"][0]["error"].lower()


class TestTableOpsRequiredParams:
    """Validate required params for table operations."""

    def test_create_table_missing_table_name(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps(
                [{"op": "create_table", "sheet": "Sheet1", "cell_ref": "A1:B2"}]
            ),
        )
        assert result["results"][0]["success"] is False
        assert "required for create_table" in result["results"][0]["error"].lower()

    def test_delete_table_missing_name(self, workbook):
        result = edit(workbook, ops=json.dumps([{"op": "delete_table"}]))
        assert result["results"][0]["success"] is False
        assert "table_name required" in result["results"][0]["error"].lower()


class TestChartOpsRequiredParams:
    """Validate required params for chart operations."""

    def test_create_chart_missing_sheet(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps(
                [{"op": "create_chart", "chart_type": "bar", "data_range": "A1:B3"}]
            ),
        )
        assert result["results"][0]["success"] is False
        assert "sheet required" in result["results"][0]["error"].lower()

    def test_create_chart_missing_chart_type(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps(
                [{"op": "create_chart", "sheet": "Sheet1", "data_range": "A1:B3"}]
            ),
        )
        assert result["results"][0]["success"] is False
        assert "chart_type required" in result["results"][0]["error"].lower()

    def test_create_chart_missing_data_range(self, workbook):
        result = edit(
            workbook,
            ops=json.dumps(
                [{"op": "create_chart", "sheet": "Sheet1", "chart_type": "bar"}]
            ),
        )
        assert result["results"][0]["success"] is False
        assert "data_range required" in result["results"][0]["error"].lower()


class TestPropertyOpsRequiredParams:
    """Validate required params for property operations."""

    def test_set_property_missing_name(self, workbook):
        result = edit(
            workbook, ops=json.dumps([{"op": "set_property", "property_value": "test"}])
        )
        assert result["results"][0]["success"] is False
        assert "property_name required" in result["results"][0]["error"].lower()

    def test_set_property_missing_value(self, workbook):
        result = edit(
            workbook, ops=json.dumps([{"op": "set_property", "property_name": "title"}])
        )
        assert result["results"][0]["success"] is False
        assert "property_value required" in result["results"][0]["error"].lower()

    def test_delete_custom_property_missing_name(self, workbook):
        result = edit(workbook, ops=json.dumps([{"op": "delete_custom_property"}]))
        assert result["results"][0]["success"] is False
        assert "property_name required" in result["results"][0]["error"].lower()
