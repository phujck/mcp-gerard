"""Shared batch edit utilities for Microsoft format MCP tools.

Provides the common batch processing pattern: parse ops JSON, resolve $prev
references, normalize text, execute operations, handle save with atomic/partial modes.
"""

from __future__ import annotations

import copy
import json
import os
import re
from collections.abc import Callable
from typing import Any, Protocol


class OpResult(Protocol):
    """Protocol for per-operation result objects."""

    index: int
    op: str
    success: bool
    element_id: str
    message: str
    error: str


class EditResult(Protocol):
    """Protocol for batch edit result objects."""

    success: bool
    message: str
    total: int
    succeeded: int
    failed: int
    saved: bool

    def model_dump(self, *, exclude_none: bool = False) -> dict[str, Any]: ...


_PREV_PATTERN = re.compile(r"^\$prev\[(\d+)\]$")
_MAX_OPS = 500


def normalize_text(
    op: str, params: dict[str, Any], text_fields: dict[str, set[str]]
) -> dict[str, Any]:
    """Normalize escaped characters in text fields (\\n -> newline, \\t -> tab)."""
    fields = text_fields.get(op, set())
    for field in fields:
        if field in params and isinstance(params[field], str):
            val = params[field].lstrip()
            if not (val.startswith("[") or val.startswith("{")):
                params[field] = params[field].replace("\\n", "\n").replace("\\t", "\t")
    return params


def resolve_prev_refs(
    params: dict[str, Any],
    results: list,
    index: int,
    prev_fields: set[str],
) -> dict[str, Any]:
    """Resolve $prev[N] references in operation parameters."""
    resolved = copy.copy(params)
    for field in prev_fields:
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


def run_batch_edit(
    *,
    file_path: str,
    ops: str,
    mode: str,
    open_pkg: Callable[[str], Any],
    new_pkg: Callable[[], Any],
    apply_op: Callable[[Any, str, dict[str, Any]], dict[str, Any]],
    make_op_result: Callable[..., Any],
    make_edit_result: Callable[..., Any],
    prev_fields: set[str],
    text_fields: dict[str, set[str]],
    excluded_ops: set[str] | None = None,
) -> dict[str, Any]:
    """Execute batch operations with atomic/partial save semantics.

    Args:
        file_path: Path to the document file
        ops: JSON array of operation objects
        mode: 'atomic' or 'partial'
        open_pkg: Callable to open an existing package from file_path
        new_pkg: Callable to create a new empty package
        apply_op: Callable(pkg, op_name, params) -> dict with message/element_id
        make_op_result: Callable to construct per-operation result objects
        make_edit_result: Callable to construct the batch edit result object
        prev_fields: Set of field names that support $prev[N] references
        text_fields: Dict mapping op names to sets of field names for text normalization
        excluded_ops: Set of op names not allowed in batch mode
    """
    if not isinstance(mode, str):
        mode = "atomic"

    try:
        operations = json.loads(ops)
    except json.JSONDecodeError as e:
        return make_edit_result(
            success=False, message=f"Invalid JSON in ops: {e}"
        ).model_dump(exclude_none=True)

    if not isinstance(operations, list):
        return make_edit_result(
            success=False, message="ops must be a JSON array"
        ).model_dump(exclude_none=True)

    if len(operations) == 0:
        return make_edit_result(success=False, message="ops array is empty").model_dump(
            exclude_none=True
        )

    if len(operations) > _MAX_OPS:
        return make_edit_result(
            success=False, message=f"ops array exceeds maximum of {_MAX_OPS} operations"
        ).model_dump(exclude_none=True)

    if mode not in ("atomic", "partial"):
        return make_edit_result(
            success=False,
            message=f"Invalid mode '{mode}': must be 'atomic' or 'partial'",
        ).model_dump(exclude_none=True)

    excluded = excluded_ops or set()
    for i, op_dict in enumerate(operations):
        if not isinstance(op_dict, dict):
            return make_edit_result(
                success=False, message=f"Operation at index {i} is not an object"
            ).model_dump(exclude_none=True)
        if "op" not in op_dict:
            return make_edit_result(
                success=False, message=f"Operation at index {i} missing 'op' field"
            ).model_dump(exclude_none=True)
        if op_dict["op"] in excluded:
            return make_edit_result(
                success=False,
                message=f"Operation '{op_dict['op']}' at index {i} is not allowed in batch mode",
            ).model_dump(exclude_none=True)

    pkg = open_pkg(file_path) if os.path.exists(file_path) else new_pkg()
    results: list = []

    for i, op_dict in enumerate(operations):
        op_name = op_dict["op"]
        params = {k: v for k, v in op_dict.items() if k != "op"}

        try:
            params = resolve_prev_refs(params, results, i, prev_fields)
            params = normalize_text(op_name, params, text_fields)
            result = apply_op(pkg, op_name, params)
            results.append(
                make_op_result(
                    index=i,
                    op=op_name,
                    success=True,
                    element_id=result.get("element_id", ""),
                    message=result.get("message", ""),
                )
            )
        except Exception as e:
            results.append(
                make_op_result(
                    index=i,
                    op=op_name,
                    success=False,
                    error=str(e),
                )
            )
            if mode == "atomic":
                break

    succeeded = sum(1 for r in results if r.success)
    failed = sum(1 for r in results if not r.success)
    total = len(operations)

    should_save = (failed == 0 and succeeded > 0) if mode == "atomic" else succeeded > 0

    saved = False
    if should_save:
        try:
            pkg.save(file_path)
            saved = True
        except Exception as e:
            return make_edit_result(
                success=False,
                message=f"Operations succeeded but save failed: {e}",
                total=total,
                succeeded=succeeded,
                failed=failed,
                results=results,
                saved=False,
            ).model_dump(exclude_none=True)

    if mode == "atomic":
        if failed > 0:
            message = f"Batch failed: {failed} operation(s) failed, file unchanged"
            success = False
        else:
            message = f"Batch completed: {succeeded}/{total} operation(s) succeeded"
            success = True
    else:
        message = (
            f"Batch completed: {succeeded}/{total} succeeded, {failed}/{total} failed"
        )
        success = succeeded > 0

    return make_edit_result(
        success=success,
        message=message,
        total=total,
        succeeded=succeeded,
        failed=failed,
        results=results,
        saved=saved,
    ).model_dump(exclude_none=True)


def require(params: dict, key: str, op: str) -> Any:
    """Get required parameter or raise ValueError. Rejects None, empty str, empty collections."""
    val = params.get(key)
    if val is None or val == "" or val == [] or val == {}:
        raise ValueError(f"{key} required for {op}")
    return val


def require_any(params: dict, key: str, op: str) -> Any:
    """Get required parameter, allowing empty string / zero / False / empty collections."""
    val = params.get(key)
    if val is None:
        raise ValueError(f"{key} required for {op}")
    return val


def convert_custom_property_value(value: Any, prop_type: str) -> tuple[Any, str]:
    """Convert a custom property value string to the appropriate Python type.

    Returns (converted_value, normalized_prop_type).
    """
    if prop_type in ("int", "i4"):
        return int(value), prop_type
    elif prop_type in ("float", "r8"):
        return float(value), prop_type
    elif prop_type == "bool":
        return str(value).lower() in ("true", "1", "yes"), prop_type
    elif prop_type in ("datetime", "filetime"):
        from datetime import datetime, timezone

        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt, "datetime"
    return value, prop_type
