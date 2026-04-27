"""
Parse Asprova-style Input Instruction (PSI) CSV exports (e.g. inputinst.csv).

Column detection follows the same approach as schedule uploads in asprova_parser.py.
"""
from __future__ import annotations

import math
from typing import Any, Dict, Mapping, Optional

from core.asprova_parser import parse_datetime


def detect_input_instruction_columns(headers: list[str]) -> Dict[str, str]:
    """Map logical field names to CSV header names."""
    mapping: Dict[str, str] = {}
    headers_lower = [h.lower().strip() for h in headers]

    def find_col(*candidates: str) -> Optional[str]:
        for c in candidates:
            c_low = c.lower()
            for i, h in enumerate(headers_lower):
                if c_low in h:
                    return headers[i]
        return None

    def find_exact(*candidates: str) -> Optional[str]:
        cand = {c.lower().strip() for c in candidates}
        for i, h in enumerate(headers_lower):
            if h in cand:
                return headers[i]
        return None

    mapping["item_code"] = find_col("inputworkinst_itemcode", "itemcode", "item_code")
    mapping["inst_time"] = find_col("inputworkinst_time", "inst_time", "time")
    mapping["quantity"] = find_col("inputworkinst_qty", "qty", "quantity")
    mapping["u_quantity"] = find_col("inputworkinst_uqty", "uqty", "u_quantity")
    mapping["qty_fixed_level"] = find_col("inputworkinst_qtyfixedlevel", "qtyfixedlevel")
    mapping["qty_fixed_level_user_specified"] = find_col(
        "inputworkinst_qtyfixedlevel_userspecified",
        "qtyfixedlevel_userspecified",
        "qtyfixedlevel_user",
    )
    mapping["pegging_method"] = find_col("inputworkinst_peggingmethod", "peggingmethod")
    mapping["object_id"] = find_exact("objectid") or find_col("object_id", "objectid")
    mapping["object_status_flag_ext"] = find_col("object_statusflagext", "statusflagext")
    mapping["flag_date"] = find_col("flagdate", "flag_date")
    mapping["operation_code"] = find_col(
        "inputworkinst_operationcode",
        "operationcode",
        "operation_code",
    )
    return mapping


def _get_val(mapping: Mapping[str, str], row: Mapping[str, Any], key: str) -> str:
    col = mapping.get(key)
    if not col:
        return ""
    v = row.get(col, "")
    return str(v).strip() if v is not None else ""


def _optional_float(raw: str) -> Optional[float]:
    if raw is None or str(raw).strip() == "":
        return None
    try:
        x = float(str(raw).strip())
        if math.isnan(x):
            return None
        return x
    except (TypeError, ValueError):
        return None


def parse_input_instruction_row(
    row: Mapping[str, Any], mapping: Mapping[str, str]
) -> Optional[Dict[str, Any]]:
    """
    Parse one CSV row into values for psi_input_instructions.
    Returns None if required fields are missing or time is invalid.
    """
    item_code = _get_val(mapping, row, "item_code")
    time_raw = _get_val(mapping, row, "inst_time")
    if not item_code or not time_raw:
        return None
    inst_dt = parse_datetime(time_raw)
    if not inst_dt:
        return None
    inst_time = inst_dt.strftime("%Y-%m-%d %H:%M:%S")

    qty = _optional_float(_get_val(mapping, row, "quantity"))
    if qty is None:
        qty = 0.0

    return {
        "item_code": item_code,
        "inst_time": inst_time,
        "quantity": qty,
        "u_quantity": _optional_float(_get_val(mapping, row, "u_quantity")),
        "qty_fixed_level": _optional_float(_get_val(mapping, row, "qty_fixed_level")),
        "qty_fixed_level_user_specified": _get_val(mapping, row, "qty_fixed_level_user_specified")
        or None,
        "pegging_method": _get_val(mapping, row, "pegging_method") or None,
        "object_id": _get_val(mapping, row, "object_id") or None,
        "object_status_flag_ext": _get_val(mapping, row, "object_status_flag_ext") or None,
        "flag_date": _get_val(mapping, row, "flag_date") or None,
        "operation_code": _get_val(mapping, row, "operation_code") or None,
    }
