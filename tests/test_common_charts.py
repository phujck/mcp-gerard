"""Tests for shared chart XML builders and embedding utilities."""

from __future__ import annotations

import io

import pytest
from lxml import etree

from mcp_gerard.microsoft.common.charts import (
    CHART_NSMAP,
    SERIES_COLORS,
    _col_letter,
    compute_chart_refs,
    create_chart_xml,
)
from mcp_gerard.microsoft.excel.embedding import create_embedded_excel
from mcp_gerard.microsoft.excel.package import ExcelPackage


class TestColLetter:
    """Tests for _col_letter column index conversion."""

    def test_single_letters(self):
        assert _col_letter(0) == "A"
        assert _col_letter(1) == "B"
        assert _col_letter(25) == "Z"

    def test_double_letters(self):
        assert _col_letter(26) == "AA"
        assert _col_letter(27) == "AB"
        assert _col_letter(51) == "AZ"
        assert _col_letter(52) == "BA"


class TestComputeChartRefs:
    """Tests for compute_chart_refs dimension-to-range conversion."""

    def test_single_series(self):
        cat_range, series = compute_chart_refs("Sheet1", n_rows=4, n_cols=2)
        assert cat_range == "'Sheet1'!$A$2:$A$4"
        assert len(series) == 1
        name_ref, name_text, val_range = series[0]
        assert name_ref == "'Sheet1'!$B$1"
        assert val_range == "'Sheet1'!$B$2:$B$4"

    def test_multi_series(self):
        cat_range, series = compute_chart_refs("Data", n_rows=5, n_cols=4)
        assert cat_range == "'Data'!$A$2:$A$5"
        assert len(series) == 3
        # Series B, C, D
        assert series[0][0] == "'Data'!$B$1"
        assert series[0][2] == "'Data'!$B$2:$B$5"
        assert series[1][0] == "'Data'!$C$1"
        assert series[2][0] == "'Data'!$D$1"

    def test_single_data_row(self):
        cat_range, series = compute_chart_refs("S", n_rows=2, n_cols=2)
        assert cat_range == "'S'!$A$2:$A$2"
        assert len(series) == 1
        assert series[0][2] == "'S'!$B$2:$B$2"


class TestCreateChartXml:
    """Tests for the shared chart XML builder."""

    def _ns(self, tag: str) -> str:
        return f"{{{CHART_NSMAP['c']}}}{tag}"

    def test_bar_chart(self):
        xml = create_chart_xml(
            "bar",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
            title="Sales",
        )
        chart = xml.find(self._ns("chart"))
        assert chart is not None
        plot = chart.find(self._ns("plotArea"))
        bar = plot.find(self._ns("barChart"))
        assert bar is not None
        bar_dir = bar.find(self._ns("barDir"))
        assert bar_dir.get("val") == "bar"

    def test_column_chart(self):
        xml = create_chart_xml(
            "column",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        bar = plot.find(self._ns("barChart"))
        assert bar is not None
        bar_dir = bar.find(self._ns("barDir"))
        assert bar_dir.get("val") == "col"

    def test_line_chart(self):
        xml = create_chart_xml(
            "line",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        assert plot.find(self._ns("lineChart")) is not None

    def test_pie_chart(self):
        xml = create_chart_xml(
            "pie",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        assert plot.find(self._ns("pieChart")) is not None

    def test_scatter_chart_has_xval_yval(self):
        xml = create_chart_xml(
            "scatter",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        scatter = plot.find(self._ns("scatterChart"))
        assert scatter is not None
        ser = scatter.find(self._ns("ser"))
        assert ser.find(self._ns("xVal")) is not None
        assert ser.find(self._ns("yVal")) is not None

    def test_area_chart(self):
        xml = create_chart_xml(
            "area",
            "Sheet1",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        assert plot.find(self._ns("areaChart")) is not None

    def test_invalid_chart_type(self):
        with pytest.raises(ValueError, match="Unsupported chart type"):
            create_chart_xml("radar", "S", "r", [("", "", "r")])

    def test_title_present(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
            title="Revenue",
        )
        chart = xml.find(self._ns("chart"))
        title_elem = chart.find(self._ns("title"))
        assert title_elem is not None
        # autoTitleDeleted should be 0
        atd = chart.find(self._ns("autoTitleDeleted"))
        assert atd.get("val") == "0"

    def test_no_title_sets_autotitledeleted(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        title_elem = chart.find(self._ns("title"))
        assert title_elem is None
        atd = chart.find(self._ns("autoTitleDeleted"))
        assert atd.get("val") == "1"

    def test_external_data_rid(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
            external_data_rid="rId1",
        )
        ext_data = xml.find(self._ns("externalData"))
        assert ext_data is not None
        r_ns = CHART_NSMAP["r"]
        assert ext_data.get(f"{{{r_ns}}}id") == "rId1"
        auto = ext_data.find(self._ns("autoUpdate"))
        assert auto.get("val") == "0"

    def test_no_external_data_without_rid(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        assert xml.find(self._ns("externalData")) is None

    def test_multi_series(self):
        xml = create_chart_xml(
            "bar",
            "S",
            "'S'!$A$2:$A$3",
            [
                ("'S'!$B$1", "S1", "'S'!$B$2:$B$3"),
                ("'S'!$C$1", "S2", "'S'!$C$2:$C$3"),
            ],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        bar = plot.find(self._ns("barChart"))
        sers = bar.findall(self._ns("ser"))
        assert len(sers) == 2
        # Check idx/order
        assert sers[0].find(self._ns("idx")).get("val") == "0"
        assert sers[1].find(self._ns("idx")).get("val") == "1"

    def test_categories_none_skips_cat(self):
        """Excel simple mode: no categories_range, no c:cat elements."""
        xml = create_chart_xml(
            "column",
            "S",
            categories_range=None,
            series=[("", "", "'S'!$A$1:$A$5")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        bar = plot.find(self._ns("barChart"))
        ser = bar.find(self._ns("ser"))
        # Should have val but NOT cat
        assert ser.find(self._ns("val")) is not None
        assert ser.find(self._ns("cat")) is None

    def test_scatter_no_categories_skips_xval(self):
        """Excel simple scatter: no xVal, only yVal."""
        xml = create_chart_xml(
            "scatter",
            "S",
            categories_range=None,
            series=[("", "", "'S'!$A$1:$A$5")],
        )
        chart = xml.find(self._ns("chart"))
        scatter = chart.find(self._ns("plotArea")).find(self._ns("scatterChart"))
        ser = scatter.find(self._ns("ser"))
        assert ser.find(self._ns("xVal")) is None
        assert ser.find(self._ns("yVal")) is not None

    def test_series_name_ref_creates_sertx(self):
        """When name_ref is provided, c:serTx is emitted."""
        xml = create_chart_xml(
            "line",
            "S",
            "'S'!$A$2:$A$3",
            [("'S'!$B$1", "Revenue", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        line = chart.find(self._ns("plotArea")).find(self._ns("lineChart"))
        ser = line.find(self._ns("ser"))
        tx = ser.find(self._ns("tx"))
        assert tx is not None

    def test_empty_name_ref_no_sertx(self):
        """When name_ref is empty string, c:serTx is NOT emitted."""
        xml = create_chart_xml(
            "line",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        line = chart.find(self._ns("plotArea")).find(self._ns("lineChart"))
        ser = line.find(self._ns("ser"))
        tx = ser.find(self._ns("tx"))
        assert tx is None

    def test_pie_chart_no_axes(self):
        """Pie charts have no axes."""
        xml = create_chart_xml(
            "pie",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        assert plot.find(self._ns("catAx")) is None
        assert plot.find(self._ns("valAx")) is None

    def test_non_pie_chart_has_axes(self):
        """Non-pie charts have category and value axes."""
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        assert plot.find(self._ns("catAx")) is not None
        assert plot.find(self._ns("valAx")) is not None

    def test_scatter_has_two_valax(self):
        """Scatter charts use valAx for both axes."""
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        plot = chart.find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        assert len(val_axes) == 2
        assert plot.find(self._ns("catAx")) is None

    def test_print_settings_present(self):
        xml = create_chart_xml(
            "column",
            "S",
            None,
            [("", "", "'S'!$A$1:$A$5")],
        )
        assert xml.find(self._ns("printSettings")) is not None

    def test_legend_present(self):
        xml = create_chart_xml(
            "column",
            "S",
            None,
            [("", "", "'S'!$A$1:$A$5")],
        )
        chart = xml.find(self._ns("chart"))
        assert chart.find(self._ns("legend")) is not None


class TestCreateEmbeddedExcel:
    """Tests for the embedded Excel workbook creation utility."""

    def test_basic_creation(self):
        data = [["Cat", "S1", "S2"], ["A", 10, 30], ["B", 20, 40]]
        xlsx_bytes, sheet_name, n_rows, n_cols = create_embedded_excel(data)
        assert isinstance(xlsx_bytes, bytes)
        assert len(xlsx_bytes) > 0
        assert sheet_name == "Sheet1"
        assert n_rows == 3
        assert n_cols == 3

    def test_single_column(self):
        data = [["Value"], [100], [200], [300]]
        xlsx_bytes, sheet_name, n_rows, n_cols = create_embedded_excel(data)
        assert n_rows == 4
        assert n_cols == 1

    def test_default_sheet_name(self):
        data = [["X"], [1]]
        xlsx_bytes, sheet_name, n_rows, n_cols = create_embedded_excel(data)
        assert sheet_name == "Sheet1"

    def test_valid_xlsx(self):
        """The returned bytes can be opened as a valid Excel package."""
        data = [["Cat", "S1"], ["A", 10], ["B", 20]]
        xlsx_bytes, sheet_name, n_rows, n_cols = create_embedded_excel(data)

        pkg = ExcelPackage.open(io.BytesIO(xlsx_bytes))
        sheet_names = [name for name, _rId, _path in pkg.get_sheet_paths()]
        assert sheet_name in sheet_names

    def test_data_roundtrip(self):
        """Values written are readable from the resulting workbook."""
        data = [["Name", "Score"], ["Alice", 95], ["Bob", 87]]
        xlsx_bytes, sheet_name, n_rows, n_cols = create_embedded_excel(data)

        from mcp_gerard.microsoft.excel.ops.cells import get_cell_value

        pkg = ExcelPackage.open(io.BytesIO(xlsx_bytes))
        assert get_cell_value(pkg, sheet_name, "A1") == "Name"
        assert get_cell_value(pkg, sheet_name, "B1") == "Score"
        assert get_cell_value(pkg, sheet_name, "A2") == "Alice"
        assert get_cell_value(pkg, sheet_name, "B2") == 95


class TestCreateChartXmlValidation:
    """Tests for input validation in create_chart_xml."""

    def test_invalid_chart_type(self):
        with pytest.raises(ValueError, match="Unsupported chart type"):
            create_chart_xml("radar", "S", "r", [("", "", "r")])

    def test_empty_series(self):
        with pytest.raises(ValueError, match="At least one data series"):
            create_chart_xml("column", "S", "'S'!$A$2:$A$3", [])


class TestScatterAxes:
    """Tests for scatter chart axes (two valAx, no catAx properties)."""

    def _ns(self, tag: str) -> str:
        return f"{{{CHART_NSMAP['c']}}}{tag}"

    def test_scatter_no_catax(self):
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        assert plot.find(self._ns("catAx")) is None

    def test_scatter_two_valax(self):
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        assert len(val_axes) == 2

    def test_scatter_x_axis_position_bottom(self):
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        # First valAx (X axis) should be at bottom
        ax_pos = val_axes[0].find(self._ns("axPos"))
        assert ax_pos.get("val") == "b"

    def test_scatter_y_axis_position_left(self):
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        # Second valAx (Y axis) should be at left
        ax_pos = val_axes[1].find(self._ns("axPos"))
        assert ax_pos.get("val") == "l"

    def test_scatter_axes_no_catax_properties(self):
        """Scatter valAx should NOT have catAx-only properties: auto, lblAlgn, lblOffset."""
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        for val_ax in plot.findall(self._ns("valAx")):
            assert val_ax.find(self._ns("auto")) is None
            assert val_ax.find(self._ns("lblAlgn")) is None
            assert val_ax.find(self._ns("lblOffset")) is None

    def test_scatter_axes_have_numfmt(self):
        """Both scatter valAx should have numFmt (value axis property)."""
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        for val_ax in plot.findall(self._ns("valAx")):
            assert val_ax.find(self._ns("numFmt")) is not None

    def test_non_scatter_catax_has_catax_properties(self):
        """Non-scatter catAx should have auto, lblAlgn, lblOffset."""
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        cat_ax = plot.find(self._ns("catAx"))
        assert cat_ax.find(self._ns("auto")) is not None
        assert cat_ax.find(self._ns("lblAlgn")) is not None
        assert cat_ax.find(self._ns("lblOffset")) is not None


class TestChartSpaceNamespace:
    """Test that chartSpace uses prefixed 'c:' namespace."""

    def test_chartspace_uses_c_prefix(self):
        xml = create_chart_xml(
            "column",
            "S",
            None,
            [("", "", "'S'!$A$1:$A$5")],
        )
        serialized = etree.tostring(xml, encoding="unicode")
        # Should use c: prefix, not default namespace
        assert "c:chartSpace" in serialized
        assert "c:chart" in serialized


class TestSeriesShapeProperties:
    """Tests for explicit c:spPr on chart series (LibreOffice compatibility)."""

    C_NS = CHART_NSMAP["c"]
    A_NS = CHART_NSMAP["a"]

    def _ns_c(self, tag: str) -> str:
        return f"{{{self.C_NS}}}{tag}"

    def _ns_a(self, tag: str) -> str:
        return f"{{{self.A_NS}}}{tag}"

    def _get_series(self, chart_type: str, n_series: int = 1):
        """Create chart and return list of c:ser elements."""
        series = [
            (f"'S'!${chr(66 + i)}$1", f"S{i}", f"'S'!${chr(66 + i)}$2:${chr(66 + i)}$3")
            for i in range(n_series)
        ]
        xml = create_chart_xml(chart_type, "S", "'S'!$A$2:$A$3", series)
        chart = xml.find(self._ns_c("chart"))
        plot = chart.find(self._ns_c("plotArea"))
        # Find the chart-type element (barChart, lineChart, etc.)
        type_map = {
            "bar": "barChart",
            "column": "barChart",
            "line": "lineChart",
            "pie": "pieChart",
            "scatter": "scatterChart",
            "area": "areaChart",
        }
        chart_elem = plot.find(self._ns_c(type_map[chart_type]))
        return chart_elem.findall(self._ns_c("ser"))

    def test_bar_series_has_solid_fill(self):
        sers = self._get_series("bar")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        assert sp_pr is not None
        solid_fill = sp_pr.find(self._ns_a("solidFill"))
        assert solid_fill is not None
        srgb_clr = solid_fill.find(self._ns_a("srgbClr"))
        assert srgb_clr is not None
        assert srgb_clr.get("val") == "4472C4"

    def test_column_series_has_solid_fill(self):
        sers = self._get_series("column")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        assert sp_pr is not None
        assert sp_pr.find(self._ns_a("solidFill")) is not None

    def test_area_series_has_solid_fill(self):
        sers = self._get_series("area")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        assert sp_pr is not None
        assert sp_pr.find(self._ns_a("solidFill")) is not None

    def test_line_series_has_no_fill_and_stroke(self):
        sers = self._get_series("line")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        assert sp_pr is not None
        # Explicit noFill
        assert sp_pr.find(self._ns_a("noFill")) is not None
        # Has line stroke
        ln = sp_pr.find(self._ns_a("ln"))
        assert ln is not None
        assert ln.get("w") == "28575"
        ln_fill = ln.find(self._ns_a("solidFill"))
        assert ln_fill is not None

    def test_scatter_series_has_no_fill_and_stroke(self):
        sers = self._get_series("scatter")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        assert sp_pr is not None
        assert sp_pr.find(self._ns_a("noFill")) is not None
        ln = sp_pr.find(self._ns_a("ln"))
        assert ln is not None

    def test_multi_series_cycles_colors(self):
        sers = self._get_series("bar", n_series=3)
        for i, ser in enumerate(sers):
            sp_pr = ser.find(self._ns_c("spPr"))
            srgb_clr = sp_pr.find(self._ns_a("solidFill")).find(self._ns_a("srgbClr"))
            expected = SERIES_COLORS[i % len(SERIES_COLORS)]
            assert srgb_clr.get("val") == expected

    def test_bar_series_has_outline(self):
        sers = self._get_series("bar")
        sp_pr = sers[0].find(self._ns_c("spPr"))
        ln = sp_pr.find(self._ns_a("ln"))
        assert ln is not None
        ln_fill = ln.find(self._ns_a("solidFill"))
        assert ln_fill is not None


class TestPieDataPointFills:
    """Tests for per-data-point fills on pie charts."""

    C_NS = CHART_NSMAP["c"]
    A_NS = CHART_NSMAP["a"]

    def _ns_c(self, tag: str) -> str:
        return f"{{{self.C_NS}}}{tag}"

    def _ns_a(self, tag: str) -> str:
        return f"{{{self.A_NS}}}{tag}"

    def test_pie_with_n_categories_has_dpt(self):
        xml = create_chart_xml(
            "pie",
            "S",
            "'S'!$A$2:$A$4",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$4")],
            n_categories=3,
        )
        chart = xml.find(self._ns_c("chart"))
        pie = chart.find(self._ns_c("plotArea")).find(self._ns_c("pieChart"))
        ser = pie.find(self._ns_c("ser"))
        dpts = ser.findall(self._ns_c("dPt"))
        assert len(dpts) == 3

    def test_pie_dpt_has_idx_and_fill(self):
        xml = create_chart_xml(
            "pie",
            "S",
            "'S'!$A$2:$A$4",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$4")],
            n_categories=3,
        )
        chart = xml.find(self._ns_c("chart"))
        pie = chart.find(self._ns_c("plotArea")).find(self._ns_c("pieChart"))
        ser = pie.find(self._ns_c("ser"))
        dpts = ser.findall(self._ns_c("dPt"))
        for i, dpt in enumerate(dpts):
            idx = dpt.find(self._ns_c("idx"))
            assert idx is not None
            assert idx.get("val") == str(i)
            sp_pr = dpt.find(self._ns_c("spPr"))
            assert sp_pr is not None
            srgb_clr = sp_pr.find(self._ns_a("solidFill")).find(self._ns_a("srgbClr"))
            expected = SERIES_COLORS[i % len(SERIES_COLORS)]
            assert srgb_clr.get("val") == expected

    def test_pie_without_n_categories_no_dpt(self):
        xml = create_chart_xml(
            "pie",
            "S",
            "'S'!$A$2:$A$4",
            [("'S'!$B$1", "S1", "'S'!$B$2:$B$4")],
        )
        chart = xml.find(self._ns_c("chart"))
        pie = chart.find(self._ns_c("plotArea")).find(self._ns_c("pieChart"))
        ser = pie.find(self._ns_c("ser"))
        dpts = ser.findall(self._ns_c("dPt"))
        assert len(dpts) == 0


class TestBarChartAxisPositions:
    """Tests for horizontal bar vs column axis positions."""

    C_NS = CHART_NSMAP["c"]

    def _ns(self, tag: str) -> str:
        return f"{{{self.C_NS}}}{tag}"

    def test_bar_chart_catax_on_left(self):
        xml = create_chart_xml(
            "bar",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        cat_ax = plot.find(self._ns("catAx"))
        assert cat_ax.find(self._ns("axPos")).get("val") == "l"

    def test_bar_chart_valax_on_bottom(self):
        xml = create_chart_xml(
            "bar",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_ax = plot.find(self._ns("valAx"))
        assert val_ax.find(self._ns("axPos")).get("val") == "b"

    def test_bar_chart_valax_has_gridlines(self):
        xml = create_chart_xml(
            "bar",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_ax = plot.find(self._ns("valAx"))
        assert val_ax.find(self._ns("majorGridlines")) is not None

    def test_column_chart_catax_on_bottom(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        cat_ax = plot.find(self._ns("catAx"))
        assert cat_ax.find(self._ns("axPos")).get("val") == "b"

    def test_column_chart_valax_on_left(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_ax = plot.find(self._ns("valAx"))
        assert val_ax.find(self._ns("axPos")).get("val") == "l"

    def test_column_chart_valax_has_gridlines(self):
        xml = create_chart_xml(
            "column",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_ax = plot.find(self._ns("valAx"))
        assert val_ax.find(self._ns("majorGridlines")) is not None

    def test_scatter_x_axis_no_gridlines(self):
        """Scatter X axis (bottom) should not have gridlines."""
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        # First valAx is X (bottom) — no gridlines
        x_ax = val_axes[0]
        assert x_ax.find(self._ns("axPos")).get("val") == "b"
        assert x_ax.find(self._ns("majorGridlines")) is None

    def test_scatter_y_axis_has_gridlines(self):
        """Scatter Y axis (left) should have gridlines."""
        xml = create_chart_xml(
            "scatter",
            "S",
            "'S'!$A$2:$A$3",
            [("", "", "'S'!$B$2:$B$3")],
        )
        plot = xml.find(self._ns("chart")).find(self._ns("plotArea"))
        val_axes = plot.findall(self._ns("valAx"))
        # Second valAx is Y (left) — has gridlines
        y_ax = val_axes[1]
        assert y_ax.find(self._ns("axPos")).get("val") == "l"
        assert y_ax.find(self._ns("majorGridlines")) is not None


class TestInlineDataCaches:
    """Tests for inline strCache/numCache elements (LibreOffice compatibility)."""

    C_NS = CHART_NSMAP["c"]
    A_NS = CHART_NSMAP["a"]

    SAMPLE_DATA = [["Cat", "Sales", "Profit"], ["Q1", 100, 40], ["Q2", 150, 60]]

    def _ns(self, tag: str) -> str:
        return f"{{{self.C_NS}}}{tag}"

    def _make_chart(self, chart_type: str, data=None, **kwargs):
        return create_chart_xml(
            chart_type,
            "Sheet1",
            "'Sheet1'!$A$2:$A$3",
            [
                ("'Sheet1'!$B$1", "Sales", "'Sheet1'!$B$2:$B$3"),
                ("'Sheet1'!$C$1", "Profit", "'Sheet1'!$C$2:$C$3"),
            ],
            data=data if data is not None else self.SAMPLE_DATA,
            **kwargs,
        )

    def test_column_has_str_cache_in_categories(self):
        xml = self._make_chart("column")
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        ser = bar.find(self._ns("ser"))
        cat = ser.find(self._ns("cat"))
        str_ref = cat.find(self._ns("strRef"))
        str_cache = str_ref.find(self._ns("strCache"))
        assert str_cache is not None
        pts = str_cache.findall(self._ns("pt"))
        assert len(pts) == 2
        assert pts[0].find(self._ns("v")).text == "Q1"
        assert pts[1].find(self._ns("v")).text == "Q2"

    def test_column_has_num_cache_in_values(self):
        xml = self._make_chart("column")
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        ser = bar.find(self._ns("ser"))
        val = ser.find(self._ns("val"))
        num_ref = val.find(self._ns("numRef"))
        num_cache = num_ref.find(self._ns("numCache"))
        assert num_cache is not None
        pts = num_cache.findall(self._ns("pt"))
        assert len(pts) == 2
        assert pts[0].find(self._ns("v")).text == "100"
        assert pts[1].find(self._ns("v")).text == "150"

    def test_series_name_has_str_cache(self):
        xml = self._make_chart("column")
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        ser = bar.find(self._ns("ser"))
        tx = ser.find(self._ns("tx"))
        str_ref = tx.find(self._ns("strRef"))
        str_cache = str_ref.find(self._ns("strCache"))
        assert str_cache is not None
        pts = str_cache.findall(self._ns("pt"))
        assert len(pts) == 1
        assert pts[0].find(self._ns("v")).text == "Sales"

    def test_no_caches_when_data_none(self):
        xml = create_chart_xml(
            "column",
            "Sheet1",
            "'Sheet1'!$A$2:$A$3",
            [("'Sheet1'!$B$1", "S1", "'Sheet1'!$B$2:$B$3")],
        )
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        ser = bar.find(self._ns("ser"))
        cat = ser.find(self._ns("cat"))
        str_ref = cat.find(self._ns("strRef"))
        assert str_ref.find(self._ns("strCache")) is None
        val = ser.find(self._ns("val"))
        num_ref = val.find(self._ns("numRef"))
        assert num_ref.find(self._ns("numCache")) is None

    def test_scatter_has_num_cache_in_xval_and_yval(self):
        scatter_data = [["X", "Y1"], [1, 10], [2, 20], [3, 30]]
        xml = create_chart_xml(
            "scatter",
            "Sheet1",
            "'Sheet1'!$A$2:$A$4",
            [("'Sheet1'!$B$1", "Y1", "'Sheet1'!$B$2:$B$4")],
            data=scatter_data,
        )
        chart = xml.find(self._ns("chart"))
        scatter = chart.find(self._ns("plotArea")).find(self._ns("scatterChart"))
        ser = scatter.find(self._ns("ser"))
        # xVal
        x_val = ser.find(self._ns("xVal"))
        x_num_ref = x_val.find(self._ns("numRef"))
        x_cache = x_num_ref.find(self._ns("numCache"))
        assert x_cache is not None
        x_pts = x_cache.findall(self._ns("pt"))
        assert len(x_pts) == 3
        assert x_pts[0].find(self._ns("v")).text == "1"
        # yVal
        y_val = ser.find(self._ns("yVal"))
        y_num_ref = y_val.find(self._ns("numRef"))
        y_cache = y_num_ref.find(self._ns("numCache"))
        assert y_cache is not None
        y_pts = y_cache.findall(self._ns("pt"))
        assert len(y_pts) == 3
        assert y_pts[0].find(self._ns("v")).text == "10"

    def test_pie_has_caches(self):
        pie_data = [["Cat", "Val"], ["A", 30], ["B", 70]]
        xml = create_chart_xml(
            "pie",
            "Sheet1",
            "'Sheet1'!$A$2:$A$3",
            [("'Sheet1'!$B$1", "Val", "'Sheet1'!$B$2:$B$3")],
            n_categories=2,
            data=pie_data,
        )
        chart = xml.find(self._ns("chart"))
        pie = chart.find(self._ns("plotArea")).find(self._ns("pieChart"))
        ser = pie.find(self._ns("ser"))
        cat = ser.find(self._ns("cat"))
        assert cat.find(self._ns("strRef")).find(self._ns("strCache")) is not None
        val = ser.find(self._ns("val"))
        assert val.find(self._ns("numRef")).find(self._ns("numCache")) is not None

    def test_all_chart_types_accept_data(self):
        for ct in ("bar", "column", "line", "pie", "scatter", "area"):
            xml = create_chart_xml(
                ct,
                "Sheet1",
                "'Sheet1'!$A$2:$A$3",
                [("'Sheet1'!$B$1", "S1", "'Sheet1'!$B$2:$B$3")],
                data=[["Cat", "S1"], ["A", 10], ["B", 20]],
            )
            assert xml is not None

    def test_ragged_row_does_not_crash(self):
        ragged_data = [["Cat", "S1", "S2"], ["A", 10, 20], ["B", 30]]
        xml = create_chart_xml(
            "column",
            "Sheet1",
            "'Sheet1'!$A$2:$A$3",
            [
                ("'Sheet1'!$B$1", "S1", "'Sheet1'!$B$2:$B$3"),
                ("'Sheet1'!$C$1", "S2", "'Sheet1'!$C$2:$C$3"),
            ],
            data=ragged_data,
        )
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        sers = bar.findall(self._ns("ser"))
        # Second series should have cache with only 1 point (B row missing col C)
        val = sers[1].find(self._ns("val"))
        num_ref = val.find(self._ns("numRef"))
        num_cache = num_ref.find(self._ns("numCache"))
        pt_count = num_cache.find(self._ns("ptCount"))
        # 1 non-None point (index 0 = 20), index 1 = None (skipped)
        assert pt_count.get("val") == "1"

    def test_second_series_values_correct(self):
        xml = self._make_chart("column")
        chart = xml.find(self._ns("chart"))
        bar = chart.find(self._ns("plotArea")).find(self._ns("barChart"))
        sers = bar.findall(self._ns("ser"))
        # Second series (Profit) should have 40 and 60
        val = sers[1].find(self._ns("val"))
        num_ref = val.find(self._ns("numRef"))
        num_cache = num_ref.find(self._ns("numCache"))
        pts = num_cache.findall(self._ns("pt"))
        assert pts[0].find(self._ns("v")).text == "40"
        assert pts[1].find(self._ns("v")).text == "60"
