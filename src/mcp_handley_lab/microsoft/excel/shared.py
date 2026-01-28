"""Core Excel functions for direct Python use.

Identical interface to MCP tools, usable without MCP server.
"""

from typing import Any


def read(
    file_path: str,
    scope: str = "sheets",
    sheet: str = "",
    range_ref: str = "",
    table_name: str = "",
    representation: str = "grid",
    include_types: bool = False,
    include_headers: bool = False,
    view: bool = False,
    limit: int = 1000,
) -> dict[str, Any]:
    """Read data from an Excel workbook.

    Uses progressive disclosure with scopes:
    - meta: Quick workbook overview
    - sheets: List of sheets for subsequent queries
    - cells: Cell values with grid (default), sparse, or detailed representation
    - table: Single table data by name
    - tables: List of all tables
    - styles: List of cell styles
    - conditional_formats: Conditional formatting rules
    - protection: Protection status
    - print_settings: Print configuration
    - charts: Chart list for a sheet
    - pivots: Pivot table list for a sheet

    Args:
        file_path: Path to .xlsx file.
        scope: What to read: meta, sheets, cells, table, tables, styles,
            conditional_formats, protection, print_settings, charts, pivots.
        sheet: Sheet name (for cells scope, defaults to first sheet).
        range_ref: Range like 'A1:C10' (for cells scope, defaults to used range).
        table_name: Table name (for table scope).
        representation: Output format: grid (2D arrays), sparse (cell list), cells (verbose).
        include_types: Include type codes (n=number, s=string, b=bool, e=error, f=formula).
        include_headers: Include header row in table data (for table scope).
        view: Include markdown table view for readability.
        limit: Maximum cells to return.

    Returns:
        Dict with scope-specific data.
    """
    from mcp_handley_lab.microsoft.excel.package import ExcelPackage
    from mcp_handley_lab.microsoft.excel.tool import (
        _read_cells,
        _read_charts,
        _read_conditional_formats,
        _read_meta,
        _read_pivots,
        _read_print_settings,
        _read_properties,
        _read_protection,
        _read_sheets,
        _read_styles,
        _read_table,
        _read_tables,
    )

    pkg = ExcelPackage.open(file_path)

    if scope == "meta":
        result = _read_meta(pkg)
    elif scope == "sheets":
        result = _read_sheets(pkg)
    elif scope == "cells":
        result = _read_cells(
            pkg, sheet, range_ref, representation, include_types, view, limit
        )
    elif scope == "table":
        result = _read_table(pkg, table_name, include_headers, limit)
    elif scope == "tables":
        result = _read_tables(pkg)
    elif scope == "styles":
        result = _read_styles(pkg)
    elif scope == "conditional_formats":
        result = _read_conditional_formats(pkg, sheet)
    elif scope == "protection":
        result = _read_protection(pkg, sheet)
    elif scope == "print_settings":
        result = _read_print_settings(pkg, sheet)
    elif scope == "charts":
        result = _read_charts(pkg, sheet)
    elif scope == "pivots":
        result = _read_pivots(pkg, sheet)
    elif scope == "properties":
        result = _read_properties(pkg)
    else:
        raise ValueError(f"Unknown scope: {scope}")

    return result.model_dump(exclude_none=True)


def edit(
    file_path: str,
    ops: str,
    mode: str = "atomic",
) -> dict[str, Any]:
    """Edit an Excel workbook using batch operations.

    Batch operations allow multiple edits in a single call with $prev chaining.
    Use read() first to discover sheets, cells, and tables.

    Args:
        file_path: Path to .xlsx file.
        ops: JSON array of operation objects, e.g.:
            [{"op": "set_cell", "sheet": "Sheet1", "cell_ref": "A1", "value": "Hello"},
             {"op": "set_style", "sheet": "Sheet1", "cell_ref": "$prev[0]", "style_index": 1}]
        mode: 'atomic' (all-or-nothing) or 'partial' (save successful ops).

    Available operations:
        - set_cell, set_formula, set_range, set_style
        - insert_rows, delete_rows, insert_columns, delete_columns
        - merge_cells, unmerge_cells
        - add_sheet, rename_sheet, delete_sheet, copy_sheet
        - create_table, delete_table, add_table_row, delete_table_row
        - add_conditional_format
        - protect_sheet, unprotect_sheet, protect_workbook, unprotect_workbook
        - lock_cells, unlock_cells
        - set_print_area, clear_print_area, set_print_titles
        - set_page_margins, set_page_orientation, set_page_size
        - set_scale, set_fit_to_page, add_page_break, clear_page_breaks
        - create_chart, delete_chart, update_chart_data
        - create_pivot, delete_pivot, refresh_pivot
        - set_meta, set_custom_property, delete_custom_property

    $prev chaining:
        Reference results of previous operations using $prev[N] where N is the
        operation index (0-based). Works for: cell_ref, range_ref, sheet,
        table_name, chart_id, pivot_id.

    Returns:
        Dict with success status, counts, and per-operation results.
    """
    from mcp_handley_lab.microsoft.excel.tool import edit as _edit

    return _edit(file_path=file_path, ops=ops, mode=mode)
