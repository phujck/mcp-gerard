"""Excel MCP tool - read and edit Excel workbooks.

Uses progressive disclosure with scopes for efficient reading.
Default representation is 'grid' with values + types arrays.
"""

from __future__ import annotations

import copy
import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_handley_lab.microsoft.common.properties import (
    delete_custom_property,
    get_core_properties,
    get_custom_properties,
    set_core_properties,
    set_custom_property,
)
from mcp_handley_lab.microsoft.excel.models import (
    CellInfo,
    CustomPropertyInfo,
    DocumentProperties,
    ExcelEditResult,
    ExcelOpResult,
    ExcelReadResult,
    GridData,
    RangeMeta,
    SheetInfo,
    SparseCell,
    TableInfo,
    WorkbookMeta,
)
from mcp_handley_lab.microsoft.excel.ops.cells import (
    get_cells_in_range,
    set_cell_formula,
    set_cell_style,
    set_cell_value,
)
from mcp_handley_lab.microsoft.excel.ops.charts import (
    create_chart,
    delete_chart,
    list_charts,
    update_chart_data,
)
from mcp_handley_lab.microsoft.excel.ops.core import (
    column_letter_to_index,
    index_to_column_letter,
    make_cell_id,
    make_sheet_id,
    make_table_id,
    parse_cell_ref,
    parse_range_ref,
)
from mcp_handley_lab.microsoft.excel.ops.formatting import (
    add_conditional_format,
    get_conditional_formats,
    list_styles,
)
from mcp_handley_lab.microsoft.excel.ops.pivots import (
    create_pivot,
    delete_pivot,
    list_pivots,
    refresh_pivot,
)
from mcp_handley_lab.microsoft.excel.ops.print_settings import (
    add_column_page_break,
    add_row_page_break,
    clear_page_breaks,
    clear_print_area,
    get_fit_to_page,
    get_page_margins,
    get_page_orientation,
    get_page_size,
    get_print_area,
    get_print_titles,
    get_scale,
    list_page_breaks,
    set_fit_to_page,
    set_page_margins,
    set_page_orientation,
    set_page_size,
    set_print_area,
    set_print_titles,
    set_scale,
)
from mcp_handley_lab.microsoft.excel.ops.protection import (
    get_sheet_protection,
    get_workbook_protection,
    is_sheet_protected,
    is_workbook_protected,
    lock_cells,
    protect_sheet,
    protect_workbook,
    unlock_cells,
    unprotect_sheet,
    unprotect_workbook,
)
from mcp_handley_lab.microsoft.excel.ops.ranges import (
    delete_columns,
    delete_rows,
    insert_columns,
    insert_rows,
    merge_cells,
    set_range_values,
    unmerge_cells,
)
from mcp_handley_lab.microsoft.excel.ops.sheets import (
    add_sheet,
    copy_sheet,
    delete_sheet,
    get_used_range,
    list_sheets,
    rename_sheet,
)
from mcp_handley_lab.microsoft.excel.ops.tables import (
    add_table_row,
    create_table,
    delete_table,
    delete_table_row,
    get_table_by_name,
    get_table_data,
    list_tables,
)
from mcp_handley_lab.microsoft.excel.package import ExcelPackage

mcp = FastMCP(
    "Excel Tool",
    instructions="""Excel workbook (.xlsx) reading and editing tool.

Scopes:
- meta: Workbook metadata (sheet count, names)
- sheets: List of all sheets
- cells: Cell values in a range (default: grid representation)
- table: Single table data by name
- tables: List of all tables
- styles: List of cell styles
- conditional_formats: Conditional formatting rules for a sheet

Representation (for cells scope):
- grid: 2D arrays of values + types (default, most compact)
- sparse: List of non-empty cells (for large sparse ranges)
- cells: Detailed per-cell objects (verbose, for editing)

Add view=true to include a markdown table view for LLM readability.
""",
)


@mcp.tool()
def read(
    file_path: str = Field(description="Path to .xlsx file"),
    scope: str = Field(
        default="sheets",
        description="What to read: meta, sheets, cells, table, tables, styles, conditional_formats, protection, print_settings, charts, properties",
    ),
    sheet: str = Field(
        default="",
        description="Sheet name (for cells scope, defaults to first sheet)",
    ),
    range_ref: str = Field(
        default="",
        description="Range like 'A1:C10' (for cells scope, defaults to used range)",
    ),
    table_name: str = Field(
        default="",
        description="Table name (for table scope)",
    ),
    representation: str = Field(
        default="grid",
        description="Output format: grid (2D arrays), sparse (cell list), cells (verbose)",
    ),
    include_types: bool = Field(
        default=False,
        description="Include type codes (n=number, s=string, b=bool, e=error, f=formula)",
    ),
    include_headers: bool = Field(
        default=False,
        description="Include header row in table data (for table scope)",
    ),
    view: bool = Field(
        default=False,
        description="Include markdown table view for readability",
    ),
    limit: int = Field(
        default=1000,
        description="Maximum cells to return",
    ),
) -> dict[str, Any]:
    """Read data from an Excel workbook. Use edit to modify. Returns block IDs for targeting edits.

    Uses progressive disclosure with scopes:
    - meta: Quick workbook overview
    - sheets: List of sheets for subsequent queries
    - cells: Cell values with grid (default), sparse, or detailed representation
    - table: Single table data by name
    - tables: List of all tables
    - styles: List of cell styles
    """
    from mcp_handley_lab.microsoft.excel.shared import read as _read

    return _read(
        file_path=file_path,
        scope=scope,
        sheet=sheet,
        range_ref=range_ref,
        table_name=table_name,
        representation=representation,
        include_types=include_types,
        include_headers=include_headers,
        view=view,
        limit=limit,
    )


def _read_meta(pkg: ExcelPackage) -> ExcelReadResult:
    """Read workbook metadata."""
    sheets = list_sheets(pkg)

    return ExcelReadResult(
        scope="meta",
        meta=WorkbookMeta(
            sheet_count=len(sheets),
            sheets=[s.name for s in sheets],
        ),
    )


def _read_sheets(pkg: ExcelPackage) -> ExcelReadResult:
    """Read list of sheets."""
    sheets = list_sheets(pkg)
    # Convert to SheetInfo with content-addressed IDs
    sheet_infos = [
        SheetInfo(
            id=make_sheet_id(s.name, s.index),
            name=s.name,
            index=s.index,
        )
        for s in sheets
    ]
    return ExcelReadResult(
        scope="sheets",
        sheets=sheet_infos,
    )


def _read_table(
    pkg: ExcelPackage, table_name: str, include_headers: bool, limit: int
) -> ExcelReadResult:
    """Read a single table by name."""
    if not table_name:
        raise ValueError("table_name is required for table scope")

    info = get_table_by_name(pkg, table_name)
    data = get_table_data(pkg, table_name, include_headers=include_headers)

    # Apply limit to data rows
    if limit and len(data) > limit:
        data = data[:limit]

    # Add content-addressed ID
    info_with_id = TableInfo(
        id=make_table_id(info.name, info.ref),
        name=info.name,
        sheet=info.sheet,
        ref=info.ref,
        columns=info.columns,
        row_count=info.row_count,
    )

    return ExcelReadResult(
        scope="table",
        table=info_with_id,
        grid=GridData(values=data),
    )


def _read_tables(pkg: ExcelPackage) -> ExcelReadResult:
    """Read list of all tables."""
    tables = list_tables(pkg)
    # Add content-addressed IDs
    tables_with_ids = [
        TableInfo(
            id=make_table_id(t.name, t.ref),
            name=t.name,
            sheet=t.sheet,
            ref=t.ref,
            columns=t.columns,
            row_count=t.row_count,
        )
        for t in tables
    ]
    return ExcelReadResult(
        scope="tables",
        tables=tables_with_ids,
    )


def _read_styles(pkg: ExcelPackage) -> ExcelReadResult:
    """Read list of cell styles."""
    styles = list_styles(pkg)
    return ExcelReadResult(
        scope="styles",
        styles=styles,
    )


def _read_conditional_formats(pkg: ExcelPackage, sheet: str) -> ExcelReadResult:
    """Read conditional formatting rules for a sheet."""
    if not sheet:
        sheets = list_sheets(pkg)
        if not sheets:
            return ExcelReadResult(scope="conditional_formats", sheet=None)
        sheet = sheets[0].name

    rules = get_conditional_formats(pkg, sheet)
    return ExcelReadResult(
        scope="conditional_formats",
        sheet=sheet,
        conditional_formats=rules,
    )


def _read_protection(pkg: ExcelPackage, sheet: str) -> ExcelReadResult:
    """Read protection status for workbook and optionally a sheet."""
    workbook_protection = get_workbook_protection(pkg)
    sheet_protection = None

    if sheet:
        sheet_protection = get_sheet_protection(pkg, sheet)

    # Build protection info dict
    protection_info = {
        "workbook": {
            "protected": is_workbook_protected(pkg),
            **(workbook_protection or {}),
        }
    }

    if sheet:
        protection_info["sheet"] = {
            "name": sheet,
            "protected": is_sheet_protected(pkg, sheet),
            **(sheet_protection or {}),
        }

    return ExcelReadResult(
        scope="protection",
        sheet=sheet or None,
        protection=protection_info,
    )


def _read_print_settings(pkg: ExcelPackage, sheet: str) -> ExcelReadResult:
    """Read print settings for a sheet."""
    if not sheet:
        sheets = list_sheets(pkg)
        if not sheets:
            return ExcelReadResult(scope="print_settings", sheet=None)
        sheet = sheets[0].name

    settings: dict = {
        "print_area": get_print_area(pkg, sheet),
        "print_titles": get_print_titles(pkg, sheet),
        "page_margins": get_page_margins(pkg, sheet),
        "page_orientation": get_page_orientation(pkg, sheet),
        "page_size": get_page_size(pkg, sheet),
        "scale": get_scale(pkg, sheet),
        "fit_to_page": get_fit_to_page(pkg, sheet),
        "page_breaks": list_page_breaks(pkg, sheet),
    }

    return ExcelReadResult(
        scope="print_settings",
        sheet=sheet,
        print_settings=settings,
    )


def _read_charts(pkg: ExcelPackage, sheet: str) -> ExcelReadResult:
    """Read charts for a sheet."""
    if not sheet:
        sheets = list_sheets(pkg)
        if not sheets:
            return ExcelReadResult(scope="charts", sheet=None, charts=[])
        sheet = sheets[0].name

    chart_list = list_charts(pkg, sheet)
    return ExcelReadResult(
        scope="charts",
        sheet=sheet,
        charts=chart_list,
    )


def _read_pivots(pkg: ExcelPackage, sheet: str) -> ExcelReadResult:
    """Read pivot tables for a sheet."""
    if not sheet:
        sheets = list_sheets(pkg)
        if not sheets:
            return ExcelReadResult(scope="pivots", sheet=None, pivots=[])
        sheet = sheets[0].name

    pivot_list = list_pivots(pkg, sheet)
    return ExcelReadResult(
        scope="pivots",
        sheet=sheet,
        pivots=pivot_list,
    )


def _get_document_properties(pkg: ExcelPackage) -> DocumentProperties:
    """Get document properties from core.xml and custom.xml."""
    core = get_core_properties(pkg)
    custom = get_custom_properties(pkg)

    return DocumentProperties(
        title=core["title"],
        author=core["author"],
        subject=core["subject"],
        keywords=core["keywords"],
        category=core["category"],
        comments=core["comments"],
        created=core["created"],
        modified=core["modified"],
        revision=core["revision"],
        last_modified_by=core["last_modified_by"],
        custom_properties=[
            CustomPropertyInfo(name=p["name"], value=p["value"], type=p["type"])
            for p in custom
        ],
    )


def _read_properties(pkg: ExcelPackage) -> ExcelReadResult:
    """Read document properties."""
    return ExcelReadResult(
        scope="properties",
        properties=_get_document_properties(pkg),
    )


def _read_cells(
    pkg: ExcelPackage,
    sheet: str,
    range_ref: str,
    representation: str,
    include_types: bool,
    include_view: bool,
    limit: int,
) -> ExcelReadResult:
    """Read cells from a sheet."""
    if not sheet:
        sheets = list_sheets(pkg)
        if not sheets:
            return ExcelReadResult(scope="cells", sheet=None)
        sheet = sheets[0].name

    # Determine range
    if not range_ref:
        range_ref = get_used_range(pkg, sheet)
        if not range_ref:
            return ExcelReadResult(scope="cells", sheet=sheet)

    # Parse range
    start_ref, end_ref = parse_range_ref(range_ref)

    # Get cells (now returns: ref, value, type_code, formula)
    raw_cells = get_cells_in_range(pkg, sheet, start_ref, end_ref)

    # Apply limit
    raw_cells = raw_cells[:limit]

    # Calculate range metadata
    start_col, start_row, _, _ = parse_cell_ref(start_ref)
    end_col, end_row, _, _ = parse_cell_ref(end_ref)
    start_col_idx = column_letter_to_index(start_col)
    end_col_idx = column_letter_to_index(end_col)

    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx
        start_col, end_col = end_col, start_col
    if start_row > end_row:
        start_row, end_row = end_row, start_row

    num_rows = end_row - start_row + 1
    num_cols = end_col_idx - start_col_idx + 1

    range_meta = RangeMeta(
        ref=range_ref,
        rows=num_rows,
        cols=num_cols,
        filled=len(raw_cells),
    )

    # Build result based on representation
    result = ExcelReadResult(scope="cells", sheet=sheet, range=range_meta)

    if representation == "grid":
        result.grid = _build_grid(
            raw_cells, start_col_idx, start_row, num_rows, num_cols, include_types
        )
    elif representation == "sparse":
        result.sparse = _build_sparse(raw_cells, include_types, sheet)
    elif representation == "cells":
        result.cells = _build_cells(raw_cells, include_types, sheet)
    else:
        raise ValueError(f"Unknown representation: {representation}")

    # Optionally add markdown view
    if include_view:
        result.view = _build_markdown_view(
            raw_cells, start_col_idx, start_row, num_rows, num_cols
        )

    return result


def _build_grid(
    cells: list[tuple[str, Any, str | None, str | None]],
    start_col_idx: int,
    start_row: int,
    num_rows: int,
    num_cols: int,
    include_types: bool,
) -> GridData:
    """Build grid representation with values array and optional types."""
    values: list[list[Any]] = [[None] * num_cols for _ in range(num_rows)]
    types: list[list[str | None]] | None = (
        [[None] * num_cols for _ in range(num_rows)] if include_types else None
    )

    for cell_ref, value, type_code, _formula in cells:
        col, row, _, _ = parse_cell_ref(cell_ref)
        col_idx = column_letter_to_index(col)
        row_offset = row - start_row
        col_offset = col_idx - start_col_idx

        if 0 <= row_offset < num_rows and 0 <= col_offset < num_cols:
            values[row_offset][col_offset] = value
            if types is not None:
                types[row_offset][col_offset] = type_code

    return GridData(values=values, types=types)


def _build_sparse(
    cells: list[tuple[str, Any, str | None, str | None]],
    include_types: bool,
    sheet_name: str,
) -> list[SparseCell]:
    """Build sparse representation for non-empty cells only."""
    return [
        SparseCell(
            id=make_cell_id(sheet_name, cell_ref, value),
            ref=cell_ref,
            value=value,
            type=type_code if include_types else None,
        )
        for cell_ref, value, type_code, _formula in cells
    ]


def _build_cells(
    cells: list[tuple[str, Any, str | None, str | None]],
    include_types: bool,
    sheet_name: str,
) -> list[CellInfo]:
    """Build detailed cell representation."""
    return [
        CellInfo(
            id=make_cell_id(sheet_name, cell_ref, value),
            ref=cell_ref,
            value=value,
            type=type_code if include_types else None,
            formula=formula,
        )
        for cell_ref, value, type_code, formula in cells
    ]


def _build_markdown_view(
    cells: list[tuple[str, Any, str | None, str | None]],
    start_col_idx: int,
    start_row: int,
    num_rows: int,
    num_cols: int,
) -> str:
    """Build markdown table view for LLM readability."""
    # Build 2D array first
    grid: list[list[Any]] = [[None] * num_cols for _ in range(num_rows)]

    for cell_ref, value, _type_code, _formula in cells:
        col, row, _, _ = parse_cell_ref(cell_ref)
        col_idx = column_letter_to_index(col)
        row_offset = row - start_row
        col_offset = col_idx - start_col_idx

        if 0 <= row_offset < num_rows and 0 <= col_offset < num_cols:
            grid[row_offset][col_offset] = value

    # Build header with column letters
    col_headers = [index_to_column_letter(start_col_idx + i) for i in range(num_cols)]
    header_row = "|   | " + " | ".join(col_headers) + " |"
    separator = "|---" + "|---" * num_cols + "|"

    # Build data rows with row numbers
    data_rows = []
    for row_offset, row_data in enumerate(grid):
        row_num = start_row + row_offset
        formatted = [_format_cell_for_markdown(v) for v in row_data]
        data_rows.append(f"| {row_num} | " + " | ".join(formatted) + " |")

    return "\n".join([header_row, separator] + data_rows)


def _format_cell_for_markdown(value: Any) -> str:
    """Format a cell value for markdown table."""
    if value is None:
        return ""
    s = str(value)
    # Escape pipes and truncate long values
    s = s.replace("|", "\\|").replace("\n", " ")
    if len(s) > 50:
        s = s[:47] + "..."
    return s


# =============================================================================
# Edit Operations - Batch API
# =============================================================================

# Pattern to match $prev[N] references
_PREV_PATTERN = re.compile(r"^\$prev\[(\d+)\]$")

# Operations that cannot be used in batch mode
_EXCLUDED_OPS = {"create", "recalculate"}

# Fields that can use $prev references
_PREV_FIELDS = {
    "cell_ref",
    "range_ref",
    "sheet",
    "table_name",
    "chart_id",
    "pivot_id",
}

# Fields that should have text normalization (\\n -> \n, \\t -> \t)
_TEXT_FIELDS: dict[str, set[str]] = {
    "set_cell": {"value"},
    "add_table_row": set(),  # value is JSON, don't normalize
}

# Maximum operations per batch
_MAX_OPS = 500


def _normalize_text(op: str, params: dict[str, Any]) -> dict[str, Any]:
    """Normalize escaped characters in text fields.

    Converts \\n to newline and \\t to tab only for known text fields,
    not for JSON payloads.
    """
    fields = _TEXT_FIELDS.get(op, set())
    for field in fields:
        if field in params and isinstance(params[field], str):
            val = params[field].lstrip()
            # Don't normalize if it looks like JSON
            if not (val.startswith("[") or val.startswith("{")):
                params[field] = params[field].replace("\\n", "\n").replace("\\t", "\t")
    return params


def _resolve_prev_refs(
    params: dict[str, Any],
    results: list[ExcelOpResult],
    index: int,
) -> dict[str, Any]:
    """Resolve $prev[N] references in operation parameters.

    Args:
        params: Operation parameters (modified in place)
        results: Results from previous operations
        index: Current operation index (for validation)

    Returns:
        Modified params dict

    Raises:
        ValueError: If $prev reference is invalid
    """
    resolved = copy.copy(params)
    for field in _PREV_FIELDS:
        if field not in resolved:
            continue
        value = resolved[field]
        if not isinstance(value, str):
            continue
        match = _PREV_PATTERN.match(value)
        if match:
            ref_idx = int(match.group(1))
            if ref_idx >= index:
                raise ValueError(
                    f"$prev[{ref_idx}] at index {index}: cannot reference future operation"
                )
            if ref_idx >= len(results):
                raise ValueError(f"$prev[{ref_idx}]: index out of range")
            prev_result = results[ref_idx]
            if not prev_result.success:
                raise ValueError(f"$prev[{ref_idx}]: referenced operation failed")
            if not prev_result.element_id:
                raise ValueError(
                    f"$prev[{ref_idx}]: referenced operation has no element_id"
                )
            resolved[field] = prev_result.element_id
    return resolved


@mcp.tool()
def edit(
    file_path: str = Field(description="Path to .xlsx file"),
    ops: str = Field(
        description='JSON array of operation objects. Each object has "op" (operation name) '
        "plus operation-specific fields. Use $prev[N] to reference element_id from operation N."
    ),
    mode: str = Field(
        default="atomic",
        description="'atomic' (save only if all succeed) or 'partial' (save if any succeed)",
    ),
) -> dict[str, Any]:
    """Edit an Excel workbook using batch operations.

    Batch operations allow multiple edits in a single call with $prev chaining.
    Use read() first to discover sheets, cells, and tables.

    Args:
        file_path: Path to .xlsx file
        ops: JSON array of operation objects, e.g.:
            [{"op": "set_cell", "sheet": "Sheet1", "cell_ref": "A1", "value": "Hello"},
             {"op": "set_style", "sheet": "Sheet1", "cell_ref": "$prev[0]", "style_index": 1}]
        mode: 'atomic' (all-or-nothing) or 'partial' (save successful ops)

    Available operations:
        - set_cell: Set cell value {sheet, cell_ref, value}
        - set_formula: Set formula {sheet, cell_ref, formula}
        - set_range: Set range values {sheet, cell_ref, value} (value is JSON 2D array)
        - set_style: Apply style {sheet, cell_ref, style_index}
        - insert_rows: Insert rows {sheet, cell_ref, count}
        - delete_rows: Delete rows {sheet, cell_ref, count}
        - insert_columns: Insert columns {sheet, cell_ref, count}
        - delete_columns: Delete columns {sheet, cell_ref, count}
        - merge_cells: Merge range {sheet, cell_ref}
        - unmerge_cells: Unmerge range {sheet, cell_ref}
        - add_sheet: Add sheet {new_name}
        - rename_sheet: Rename sheet {sheet, new_name}
        - delete_sheet: Delete sheet {sheet}
        - copy_sheet: Copy sheet {sheet, new_name}
        - create_table: Create table {sheet, cell_ref, table_name}
        - delete_table: Delete table {table_name}
        - add_table_row: Add row {table_name, value} (value is JSON array)
        - delete_table_row: Delete row {table_name, row_index}
        - add_conditional_format: Add conditional format {sheet, cell_ref, rule_type, operator, formula, style_index, priority}
        - protect_sheet/unprotect_sheet: {sheet, password}
        - protect_workbook/unprotect_workbook: {password}
        - lock_cells/unlock_cells: {sheet, cell_ref}
        - set_print_area/clear_print_area: {sheet, cell_ref}
        - set_print_titles: {sheet, print_rows, print_cols}
        - set_page_margins: {sheet, margin_left, margin_right, margin_top, margin_bottom}
        - set_page_orientation: {sheet, landscape}
        - set_page_size: {sheet, paper_size}
        - set_scale: {sheet, scale}
        - set_fit_to_page: {sheet, fit_width, fit_height}
        - add_page_break: {sheet, break_type, break_position}
        - clear_page_breaks: {sheet}
        - create_chart: {sheet, chart_type, data_range, position, title}
        - delete_chart: {sheet, chart_id}
        - update_chart_data: {sheet, chart_id, data_range}
        - create_pivot: {sheet, data_range, position, row_fields, col_fields, value_fields, pivot_name, agg_func}
        - delete_pivot: {sheet, pivot_id}
        - refresh_pivot: {sheet, pivot_id}
        - set_meta: {property_name, property_value}
        - set_custom_property: {property_name, property_value, property_type}
        - delete_custom_property: {property_name}

    $prev chaining:
        Reference results of previous operations using $prev[N] where N is the
        operation index (0-based). Only works for: cell_ref, range_ref, sheet,
        table_name, chart_id, pivot_id.

    Returns:
        ExcelEditResult with success status, counts, and per-operation results
    """
    # Handle direct calls (Field descriptors not resolved outside MCP)
    if not isinstance(mode, str):
        mode = "atomic"

    # Parse operations
    try:
        operations = json.loads(ops)
    except json.JSONDecodeError as e:
        return ExcelEditResult(
            success=False,
            message=f"Invalid JSON in ops: {e}",
        ).model_dump(exclude_none=True)

    if not isinstance(operations, list):
        return ExcelEditResult(
            success=False,
            message="ops must be a JSON array",
        ).model_dump(exclude_none=True)

    if len(operations) == 0:
        return ExcelEditResult(
            success=False,
            message="ops array is empty",
        ).model_dump(exclude_none=True)

    if len(operations) > _MAX_OPS:
        return ExcelEditResult(
            success=False,
            message=f"ops array exceeds maximum of {_MAX_OPS} operations",
        ).model_dump(exclude_none=True)

    if mode not in ("atomic", "partial"):
        return ExcelEditResult(
            success=False,
            message=f"Invalid mode '{mode}': must be 'atomic' or 'partial'",
        ).model_dump(exclude_none=True)

    # Validate operations
    for i, op_dict in enumerate(operations):
        if not isinstance(op_dict, dict):
            return ExcelEditResult(
                success=False,
                message=f"Operation at index {i} is not an object",
            ).model_dump(exclude_none=True)
        if "op" not in op_dict:
            return ExcelEditResult(
                success=False,
                message=f"Operation at index {i} missing 'op' field",
            ).model_dump(exclude_none=True)
        if op_dict["op"] in _EXCLUDED_OPS:
            return ExcelEditResult(
                success=False,
                message=f"Operation '{op_dict['op']}' at index {i} is not allowed in batch mode",
            ).model_dump(exclude_none=True)

    # Open package once
    pkg = ExcelPackage.open(file_path)
    results: list[ExcelOpResult] = []

    # Execute operations
    for i, op_dict in enumerate(operations):
        op_name = op_dict["op"]
        params = {k: v for k, v in op_dict.items() if k != "op"}

        try:
            # Resolve $prev references
            params = _resolve_prev_refs(params, results, i)

            # Normalize text fields
            params = _normalize_text(op_name, params)

            # Execute operation
            result = _apply_op(pkg, op_name, params)
            results.append(
                ExcelOpResult(
                    index=i,
                    op=op_name,
                    success=True,
                    element_id=result.get("element_id", ""),
                    message=result.get("message", ""),
                )
            )
        except Exception as e:
            results.append(
                ExcelOpResult(
                    index=i,
                    op=op_name,
                    success=False,
                    error=str(e),
                )
            )
            # In atomic mode, stop on first failure
            if mode == "atomic":
                break

    # Calculate summary
    succeeded = sum(1 for r in results if r.success)
    failed = sum(1 for r in results if not r.success)
    total = len(operations)

    # Determine if we should save
    should_save = (failed == 0 and succeeded > 0) if mode == "atomic" else succeeded > 0

    # Save if appropriate
    saved = False
    if should_save:
        try:
            pkg.save(file_path)
            saved = True
        except Exception as e:
            # If save fails, report it but keep operation results
            return ExcelEditResult(
                success=False,
                message=f"Operations succeeded but save failed: {e}",
                total=total,
                succeeded=succeeded,
                failed=failed,
                results=results,
                saved=False,
            ).model_dump(exclude_none=True)

    # Build response
    if mode == "atomic":
        if failed > 0:
            message = f"Batch failed: {failed} operation(s) failed, file unchanged"
            success = False
        else:
            message = f"Batch completed: {succeeded}/{total} operation(s) succeeded"
            success = True
    else:  # partial
        message = (
            f"Batch completed: {succeeded}/{total} succeeded, {failed}/{total} failed"
        )
        success = succeeded > 0

    return ExcelEditResult(
        success=success,
        message=message,
        total=total,
        succeeded=succeeded,
        failed=failed,
        results=results,
        saved=saved,
    ).model_dump(exclude_none=True)


def _apply_op(pkg: ExcelPackage, op: str, params: dict[str, Any]) -> dict[str, Any]:
    """Apply a single operation to an open package.

    Returns dict with 'message' and optionally 'element_id' for $prev chaining.
    """
    if op == "set_cell":
        return _op_set_cell(pkg, params)
    elif op == "set_formula":
        return _op_set_formula(pkg, params)
    elif op == "set_range":
        return _op_set_range(pkg, params)
    elif op == "set_style":
        return _op_set_style(pkg, params)
    elif op == "insert_rows":
        return _op_insert_rows(pkg, params)
    elif op == "delete_rows":
        return _op_delete_rows(pkg, params)
    elif op == "insert_columns":
        return _op_insert_columns(pkg, params)
    elif op == "delete_columns":
        return _op_delete_columns(pkg, params)
    elif op == "merge_cells":
        return _op_merge_cells(pkg, params)
    elif op == "unmerge_cells":
        return _op_unmerge_cells(pkg, params)
    elif op == "add_sheet":
        return _op_add_sheet(pkg, params)
    elif op == "rename_sheet":
        return _op_rename_sheet(pkg, params)
    elif op == "delete_sheet":
        return _op_delete_sheet(pkg, params)
    elif op == "copy_sheet":
        return _op_copy_sheet(pkg, params)
    elif op == "create_table":
        return _op_create_table(pkg, params)
    elif op == "delete_table":
        return _op_delete_table(pkg, params)
    elif op == "add_table_row":
        return _op_add_table_row(pkg, params)
    elif op == "delete_table_row":
        return _op_delete_table_row(pkg, params)
    elif op == "add_conditional_format":
        return _op_add_conditional_format(pkg, params)
    elif op == "protect_sheet":
        return _op_protect_sheet(pkg, params)
    elif op == "unprotect_sheet":
        return _op_unprotect_sheet(pkg, params)
    elif op == "protect_workbook":
        return _op_protect_workbook(pkg, params)
    elif op == "unprotect_workbook":
        return _op_unprotect_workbook(pkg, params)
    elif op == "lock_cells":
        return _op_lock_cells(pkg, params)
    elif op == "unlock_cells":
        return _op_unlock_cells(pkg, params)
    elif op == "set_print_area":
        return _op_set_print_area(pkg, params)
    elif op == "clear_print_area":
        return _op_clear_print_area(pkg, params)
    elif op == "set_print_titles":
        return _op_set_print_titles(pkg, params)
    elif op == "set_page_margins":
        return _op_set_page_margins(pkg, params)
    elif op == "set_page_orientation":
        return _op_set_page_orientation(pkg, params)
    elif op == "set_page_size":
        return _op_set_page_size(pkg, params)
    elif op == "set_scale":
        return _op_set_scale(pkg, params)
    elif op == "set_fit_to_page":
        return _op_set_fit_to_page(pkg, params)
    elif op == "add_page_break":
        return _op_add_page_break(pkg, params)
    elif op == "clear_page_breaks":
        return _op_clear_page_breaks(pkg, params)
    elif op == "create_chart":
        return _op_create_chart(pkg, params)
    elif op == "delete_chart":
        return _op_delete_chart(pkg, params)
    elif op == "update_chart_data":
        return _op_update_chart_data(pkg, params)
    elif op == "create_pivot":
        return _op_create_pivot(pkg, params)
    elif op == "delete_pivot":
        return _op_delete_pivot(pkg, params)
    elif op == "refresh_pivot":
        return _op_refresh_pivot(pkg, params)
    elif op == "set_meta":
        return _op_set_meta(pkg, params)
    elif op == "set_custom_property":
        return _op_set_custom_property(pkg, params)
    elif op == "delete_custom_property":
        return _op_delete_custom_property(pkg, params)
    else:
        raise ValueError(f"Unknown operation: {op}")


# =============================================================================
# Individual Operation Handlers
# =============================================================================


def _op_set_cell(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set a cell's value."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    value = params.get("value", "")

    # Auto-detect type from value string
    parsed_value: Any = value
    if value == "":
        parsed_value = None
    elif isinstance(value, str):
        if value.lower() == "true":
            parsed_value = True
        elif value.lower() == "false":
            parsed_value = False
        else:
            try:
                parsed_value = float(value) if "." in value else int(value)
            except ValueError:
                parsed_value = value

    set_cell_value(pkg, sheet, cell_ref, parsed_value)
    return {
        "message": f"Set {cell_ref} to {repr(parsed_value)}",
        "element_id": cell_ref,
    }


def _op_set_formula(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set a cell's formula."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    formula = params.get("formula", "")

    set_cell_formula(pkg, sheet, cell_ref, formula)
    return {"message": f"Set {cell_ref} formula to ={formula}", "element_id": cell_ref}


def _op_set_range(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set range values from JSON 2D array."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    value = params.get("value", "")

    values = json.loads(value) if isinstance(value, str) else value
    set_range_values(pkg, sheet, cell_ref, values)
    return {"message": f"Set range starting at {cell_ref}", "element_id": cell_ref}


def _op_set_style(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Apply a style to a cell or range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    style_index = params.get("style_index", -1)

    if ":" in cell_ref:
        # Range - apply style to all cells
        start_ref, end_ref = parse_range_ref(cell_ref)
        start_col, start_row, _, _ = parse_cell_ref(start_ref)
        end_col, end_row, _, _ = parse_cell_ref(end_ref)
        start_col_idx = column_letter_to_index(start_col)
        end_col_idx = column_letter_to_index(end_col)

        for row_num in range(start_row, end_row + 1):
            for col_idx in range(start_col_idx, end_col_idx + 1):
                col_letter = index_to_column_letter(col_idx)
                ref = f"{col_letter}{row_num}"
                set_cell_style(pkg, sheet, ref, style_index)

        return {
            "message": f"Applied style {style_index} to range {cell_ref}",
            "element_id": cell_ref,
        }
    else:
        set_cell_style(pkg, sheet, cell_ref, style_index)
        return {
            "message": f"Applied style {style_index} to {cell_ref}",
            "element_id": cell_ref,
        }


def _op_insert_rows(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Insert rows."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    count = params.get("count", 1)

    _, row, _, _ = parse_cell_ref(cell_ref)
    insert_rows(pkg, sheet, row, count)
    return {"message": f"Inserted {count} row(s) at row {row}", "element_id": ""}


def _op_delete_rows(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete rows."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    count = params.get("count", 1)

    _, row, _, _ = parse_cell_ref(cell_ref)
    delete_rows(pkg, sheet, row, count)
    return {
        "message": f"Deleted {count} row(s) starting at row {row}",
        "element_id": "",
    }


def _op_insert_columns(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Insert columns."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    count = params.get("count", 1)

    col, _, _, _ = parse_cell_ref(cell_ref)
    col_idx = column_letter_to_index(col)
    insert_columns(pkg, sheet, col_idx, count)
    return {"message": f"Inserted {count} column(s) at column {col}", "element_id": ""}


def _op_delete_columns(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete columns."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    count = params.get("count", 1)

    col, _, _, _ = parse_cell_ref(cell_ref)
    col_idx = column_letter_to_index(col)
    delete_columns(pkg, sheet, col_idx, count)
    return {
        "message": f"Deleted {count} column(s) starting at column {col}",
        "element_id": "",
    }


def _op_merge_cells(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Merge cells in range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")

    merge_cells(pkg, sheet, cell_ref)
    return {"message": f"Merged cells {cell_ref}", "element_id": cell_ref}


def _op_unmerge_cells(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Unmerge cells in range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")

    unmerge_cells(pkg, sheet, cell_ref)
    return {"message": f"Unmerged cells {cell_ref}", "element_id": cell_ref}


def _op_add_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Add new sheet."""
    name = params.get("new_name", "") or params.get("name", "")

    add_sheet(pkg, name)
    return {"message": f"Added sheet '{name}'", "element_id": name}


def _op_rename_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Rename existing sheet."""
    old_name = params.get("sheet", "")
    new_name = params.get("new_name", "")

    rename_sheet(pkg, old_name, new_name)
    return {
        "message": f"Renamed sheet '{old_name}' to '{new_name}'",
        "element_id": new_name,
    }


def _op_delete_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete sheet."""
    name = params.get("sheet", "")

    delete_sheet(pkg, name)
    return {"message": f"Deleted sheet '{name}'", "element_id": ""}


def _op_copy_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Copy sheet to new sheet."""
    source = params.get("sheet", "")
    new_name = params.get("new_name", "")

    copy_sheet(pkg, source, new_name)
    return {
        "message": f"Copied sheet '{source}' to '{new_name}'",
        "element_id": new_name,
    }


def _op_create_table(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Create table from range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    table_name = params.get("table_name", "") or params.get("new_name", "")

    create_table(pkg, sheet, cell_ref, table_name)
    return {
        "message": f"Created table '{table_name}' from range {cell_ref}",
        "element_id": table_name,
    }


def _op_delete_table(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete table by name."""
    table_name = params.get("table_name", "")

    delete_table(pkg, table_name)
    return {"message": f"Deleted table '{table_name}'", "element_id": ""}


def _op_add_table_row(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Add row to table."""
    table_name = params.get("table_name", "")
    value = params.get("value", "")

    row_data = json.loads(value) if isinstance(value, str) else value
    add_table_row(pkg, table_name, row_data)
    return {"message": f"Added row to table '{table_name}'", "element_id": table_name}


def _op_delete_table_row(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete row from table."""
    table_name = params.get("table_name", "")
    row_index = params.get("row_index", 0)

    delete_table_row(pkg, table_name, row_index)
    return {
        "message": f"Deleted row {row_index} from table '{table_name}'",
        "element_id": table_name,
    }


def _op_add_conditional_format(
    pkg: ExcelPackage, params: dict[str, Any]
) -> dict[str, Any]:
    """Add conditional formatting."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")
    rule_type = params.get("rule_type", "")
    operator = params.get("operator", "")
    formula = params.get("formula", "")
    style_index = params.get("style_index", -1)
    priority = params.get("priority", 1)

    add_conditional_format(
        pkg, sheet, cell_ref, rule_type, operator, formula, style_index, priority
    )
    return {
        "message": f"Added conditional format to {cell_ref}",
        "element_id": cell_ref,
    }


def _op_protect_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Protect sheet."""
    sheet = params.get("sheet", "")
    password = params.get("password", "")

    protect_sheet(pkg, sheet, password if password else None)
    return {"message": f"Protected sheet '{sheet}'", "element_id": ""}


def _op_unprotect_sheet(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Unprotect sheet."""
    sheet = params.get("sheet", "")
    password = params.get("password", "")

    unprotect_sheet(pkg, sheet, password if password else None)
    return {"message": f"Unprotected sheet '{sheet}'", "element_id": ""}


def _op_protect_workbook(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Protect workbook."""
    password = params.get("password", "")

    protect_workbook(pkg, password if password else None)
    return {"message": "Protected workbook", "element_id": ""}


def _op_unprotect_workbook(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Unprotect workbook."""
    password = params.get("password", "")

    unprotect_workbook(pkg, password if password else None)
    return {"message": "Unprotected workbook", "element_id": ""}


def _op_lock_cells(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Lock cells in range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")

    lock_cells(pkg, sheet, cell_ref)
    return {"message": f"Locked cells {cell_ref}", "element_id": cell_ref}


def _op_unlock_cells(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Unlock cells in range."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")

    unlock_cells(pkg, sheet, cell_ref)
    return {"message": f"Unlocked cells {cell_ref}", "element_id": cell_ref}


def _op_set_print_area(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set print area."""
    sheet = params.get("sheet", "")
    cell_ref = params.get("cell_ref", "")

    set_print_area(pkg, sheet, cell_ref)
    return {"message": f"Set print area to {cell_ref}", "element_id": ""}


def _op_clear_print_area(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Clear print area."""
    sheet = params.get("sheet", "")

    clear_print_area(pkg, sheet)
    return {"message": f"Cleared print area for sheet '{sheet}'", "element_id": ""}


def _op_set_print_titles(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set print titles."""
    sheet = params.get("sheet", "")
    print_rows = params.get("print_rows", "")
    print_cols = params.get("print_cols", "")

    set_print_titles(pkg, sheet, print_rows or None, print_cols or None)
    return {"message": f"Set print titles for sheet '{sheet}'", "element_id": ""}


def _op_set_page_margins(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set page margins."""
    sheet = params.get("sheet", "")
    left = params.get("margin_left", -1.0)
    right = params.get("margin_right", -1.0)
    top = params.get("margin_top", -1.0)
    bottom = params.get("margin_bottom", -1.0)

    set_page_margins(
        pkg,
        sheet,
        left if left >= 0 else None,
        right if right >= 0 else None,
        top if top >= 0 else None,
        bottom if bottom >= 0 else None,
    )
    return {"message": f"Set page margins for sheet '{sheet}'", "element_id": ""}


def _op_set_page_orientation(
    pkg: ExcelPackage, params: dict[str, Any]
) -> dict[str, Any]:
    """Set page orientation."""
    sheet = params.get("sheet", "")
    landscape = params.get("landscape", False)

    set_page_orientation(pkg, sheet, landscape)
    orientation = "landscape" if landscape else "portrait"
    return {"message": f"Set {sheet} to {orientation}", "element_id": ""}


def _op_set_page_size(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set page size."""
    sheet = params.get("sheet", "")
    paper_size = params.get("paper_size", 1)

    set_page_size(pkg, sheet, paper_size)
    return {
        "message": f"Set paper size to {paper_size} for sheet '{sheet}'",
        "element_id": "",
    }


def _op_set_scale(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set print scale."""
    sheet = params.get("sheet", "")
    scale_value = params.get("scale", 100)

    set_scale(pkg, sheet, scale_value)
    return {
        "message": f"Set scale to {scale_value}% for sheet '{sheet}'",
        "element_id": "",
    }


def _op_set_fit_to_page(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set fit to page."""
    sheet = params.get("sheet", "")
    fit_width = params.get("fit_width", -1)
    fit_height = params.get("fit_height", -1)

    set_fit_to_page(
        pkg,
        sheet,
        fit_width if fit_width >= 0 else None,
        fit_height if fit_height >= 0 else None,
    )
    return {"message": f"Set fit to page for sheet '{sheet}'", "element_id": ""}


def _op_add_page_break(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Add page break."""
    sheet = params.get("sheet", "")
    break_type = params.get("break_type", "row")
    break_position = params.get("break_position", 0)

    if break_type == "row":
        add_row_page_break(pkg, sheet, break_position)
        return {
            "message": f"Added row page break at row {break_position}",
            "element_id": "",
        }
    else:
        add_column_page_break(pkg, sheet, break_position)
        return {
            "message": f"Added column page break at column {break_position}",
            "element_id": "",
        }


def _op_clear_page_breaks(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Clear all page breaks."""
    sheet = params.get("sheet", "")

    clear_page_breaks(pkg, sheet)
    return {"message": f"Cleared all page breaks for sheet '{sheet}'", "element_id": ""}


def _op_create_chart(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Create chart."""
    sheet = params.get("sheet", "")
    chart_type = params.get("chart_type", "")
    data_range = params.get("data_range", "")
    position = params.get("position", "")
    title = params.get("title", "")

    chart_id = create_chart(pkg, sheet, chart_type, data_range, position, title or None)
    return {
        "message": f"Created {chart_type} chart at {position}",
        "element_id": chart_id,
    }


def _op_delete_chart(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete chart."""
    sheet = params.get("sheet", "")
    chart_id = params.get("chart_id", "")

    delete_chart(pkg, sheet, chart_id)
    return {"message": f"Deleted chart '{chart_id}'", "element_id": ""}


def _op_update_chart_data(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Update chart data range."""
    sheet = params.get("sheet", "")
    chart_id = params.get("chart_id", "")
    data_range = params.get("data_range", "")

    update_chart_data(pkg, sheet, chart_id, data_range)
    return {
        "message": f"Updated chart '{chart_id}' data to {data_range}",
        "element_id": chart_id,
    }


def _op_create_pivot(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Create pivot table."""
    sheet = params.get("sheet", "")
    data_range = params.get("data_range", "")
    position = params.get("position", "")
    row_fields_str = params.get("row_fields", "")
    col_fields_str = params.get("col_fields", "")
    value_fields_str = params.get("value_fields", "")
    pivot_name = params.get("pivot_name", "")
    agg_func = params.get("agg_func", "sum")

    row_fields = [f.strip() for f in row_fields_str.split(",") if f.strip()]
    col_fields = [f.strip() for f in col_fields_str.split(",") if f.strip()]
    value_fields = [f.strip() for f in value_fields_str.split(",") if f.strip()]

    pivot_id = create_pivot(
        pkg,
        sheet,
        data_range,
        position,
        row_fields,
        col_fields,
        value_fields,
        pivot_name or None,
        agg_func,
    )
    return {"message": f"Created pivot table at {position}", "element_id": pivot_id}


def _op_delete_pivot(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Delete pivot table."""
    sheet = params.get("sheet", "")
    pivot_id = params.get("pivot_id", "")

    delete_pivot(pkg, sheet, pivot_id)
    return {"message": f"Deleted pivot table '{pivot_id}'", "element_id": ""}


def _op_refresh_pivot(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Refresh pivot table."""
    sheet = params.get("sheet", "")
    pivot_id = params.get("pivot_id", "")

    refresh_pivot(pkg, sheet, pivot_id)
    return {"message": f"Refreshed pivot table '{pivot_id}'", "element_id": pivot_id}


def _op_set_meta(pkg: ExcelPackage, params: dict[str, Any]) -> dict[str, Any]:
    """Set core document property."""
    name = params.get("property_name", "")
    value = params.get("property_value", "")

    set_core_properties(pkg, **{name: value})
    return {"message": f"Set core property '{name}' = '{value}'", "element_id": ""}


def _op_set_custom_property(
    pkg: ExcelPackage, params: dict[str, Any]
) -> dict[str, Any]:
    """Set custom document property."""
    from datetime import datetime, timezone

    name = params.get("property_name", "")
    value = params.get("property_value", "")
    prop_type = params.get("property_type", "string")

    # Convert value to appropriate type
    actual_value: Any = value
    if prop_type in ("int", "i4"):
        actual_value = int(value)
    elif prop_type in ("float", "r8"):
        actual_value = float(value)
    elif prop_type == "bool":
        actual_value = value.lower() in ("true", "1", "yes")
    elif prop_type in ("datetime", "filetime"):
        # Parse ISO format datetime
        prop_type = "datetime"
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        actual_value = dt

    set_custom_property(pkg, name, actual_value, prop_type)
    return {
        "message": f"Set custom property '{name}' = '{value}' ({prop_type})",
        "element_id": "",
    }


def _op_delete_custom_property(
    pkg: ExcelPackage, params: dict[str, Any]
) -> dict[str, Any]:
    """Delete custom document property."""
    name = params.get("property_name", "")

    deleted = delete_custom_property(pkg, name)
    if deleted:
        return {"message": f"Deleted custom property '{name}'", "element_id": ""}
    else:
        raise ValueError(f"Custom property '{name}' not found")


# =============================================================================
# Standalone Operations (not available in batch mode)
# =============================================================================


@mcp.tool()
def create(
    file_path: str = Field(description="Path to create new .xlsx file"),
) -> dict[str, Any]:
    """Create a new empty Excel workbook."""
    pkg = ExcelPackage.new()
    pkg.save(file_path)
    return ExcelEditResult(
        success=True,
        message=f"Created workbook: {file_path}",
        total=1,
        succeeded=1,
        saved=True,
    ).model_dump(exclude_none=True)


@mcp.tool()
def recalculate(
    file_path: str = Field(description="Path to .xlsx file"),
) -> dict[str, Any]:
    """Recalculate all formulas using LibreOffice headless.

    Opens the file in LibreOffice to trigger formula calculation,
    then saves it back. The cached values in <v> elements are then populated.
    """
    resolved_path = str(Path(file_path).resolve())
    file_name = Path(resolved_path).name

    with tempfile.TemporaryDirectory() as tmpdir:
        input_dir = Path(tmpdir) / "in"
        output_dir = Path(tmpdir) / "out"
        input_dir.mkdir()
        output_dir.mkdir()

        input_copy = input_dir / file_name
        shutil.copy2(resolved_path, input_copy)

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--calc",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(output_dir),
                str(input_copy),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice failed: {result.stderr or result.stdout}")

        output_file = output_dir / file_name
        if not output_file.exists():
            xlsx_files = list(output_dir.glob("*.xlsx"))
            if len(xlsx_files) == 1:
                output_file = xlsx_files[0]
            elif not xlsx_files:
                raise RuntimeError(
                    f"LibreOffice did not produce any .xlsx file in {output_dir}"
                )
            else:
                raise RuntimeError(
                    f"LibreOffice produced multiple files: {[f.name for f in xlsx_files]}"
                )

        shutil.move(str(output_file), resolved_path)

    return ExcelEditResult(
        success=True,
        message="Recalculated all formulas",
        total=1,
        succeeded=1,
        saved=True,
    ).model_dump(exclude_none=True)
