from __future__ import annotations

import math
from datetime import datetime
from typing import Any, Dict, Mapping, Optional


def parse_duration_minutes(raw: Any) -> Optional[float]:
    """
    Parse a duration field to minutes.
    Accepts seconds (common in exports) or minutes. Heuristic:
    - if value >= 3600 and divisible-ish by 60 -> treat as seconds
    - else treat as minutes
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "":
        return None
    try:
        v = float(s)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    # Heuristic seconds detection
    if v >= 3600 and abs((v / 60) - round(v / 60)) < 1e-6:
        return v / 60.0
    return v


def parse_datetime(s: Any) -> Optional[datetime]:
    """Try multiple datetime formats used by exports."""
    if s is None:
        return None
    if str(s).strip() == "":
        return None
    s = str(s).strip()
    formats = [
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def detect_columns(headers: list[str]) -> Dict[str, str]:
    """Auto-detect column mappings from CSV headers."""
    mapping: Dict[str, str] = {}
    headers_lower = [h.lower().strip() for h in headers]

    def find_exact(*candidates: str) -> Optional[str]:
        cand = {c.lower().strip() for c in candidates}
        for i, h in enumerate(headers_lower):
            if h in cand:
                return headers[i]
        return None

    def find_col(*candidates: str) -> Optional[str]:
        for c in candidates:
            for i, h in enumerate(headers_lower):
                if c in h:
                    return headers[i]
        return None

    # Asprova-like exports often use "Work_*" column names; include those keywords.
    mapping["order_id"] = find_col(
        "work_ordercode",
        "ordercode",
        "order",
        "lot",
        "job",
        "work order",
        "wo_",
    )
    mapping["order_item_code"] = find_col(
        "workuser_orderitem",
        "workuserorderitem",
        "work_orderitem",
        "orderitem",
        "order_item",
    )
    mapping["operation_out_item"] = find_col(
        "work_operationoutmainitem",
        "operationoutmainitem",
    )
    mapping["operation_id"] = find_col("objectid", "operationid", "op_id", "opid")

    # Distinguish ID vs Code fields to avoid substring collisions.
    mapping["operation_code"] = find_exact("work_code") or find_col(
        "work_code",
        "operationcode",
        "op_code",
        "opcode",
    )
    mapping["next_operation_code"] = find_exact("work_nextoperationcode") or find_col(
        "work_nextoperationcode",
        "nextoperationcode",
        "next_operation_code",
    )
    mapping["next_operation_id"] = find_exact("work_nextoperation") or find_col(
        "work_nextoperation_id",
        "nextoperationid",
        "next_operation_id",
        "work_nextoperation",
    )
    mapping["item_id"] = find_col(
        "work_operationoutmainitemcode",
        "outmainitemcode",
        "item_id",
        "item id",
        "part_id",
        "product_id",
        "sku",
    )
    mapping["item_name"] = find_col(
        # Avoid overly generic keywords like "product" which can match "productionStartTime".
        "item_name",
        "item name",
        "description",
        "part_name",
        "part name",
    )
    mapping["machine_id"] = find_col(
        "work_operationmainrescode",
        "work_resultmainrescode",
        "mainrescode",
        "rescode",
        "machine_id",
        "machine id",
        "equipment_id",
        "res_id",
        "resource_id",
        "resource id",
    )
    mapping["machine_name"] = find_col(
        "work_operationmainrescode",
        "work_resultmainrescode",
        "mainrescode",
        "rescode",
        "machine_name",
        "machine name",
        "resource_name",
        "resource name",
        "equipment_name",
        "equipment",
        "resource",
    )
    mapping["start_time"] = find_col(
        "work_operationproductionstarttime",
        "operationproductionstarttime",
        "work_resultstarttime",
        "resultstarttime",
        "work_uest",
        "start",
        "begin",
        "from",
        "planned_start",
        "schedule_start",
    )
    mapping["end_time"] = find_col(
        "work_operationproductionendtime",
        "operationproductionendtime",
        "work_resultendtime",
        "resultendtime",
        "work_ulet",
        "end",
        "finish",
        "to",
        "planned_end",
        "schedule_end",
        "completion",
    )
    mapping["quantity"] = find_col(
        "work_operationoutmainitemqty",
        "outmainitemqty",
        "work_resultqty",
        "resultqty",
        "qty",
        "quantity",
        "amount",
        "volume",
        "count",
    )
    mapping["status"] = find_col("work_status", "status", "state", "condition")
    mapping["process_name"] = find_col(
        "work_code",
        "work_type",
        "work_operationoutmainitemcode",
        "process",
        "operation",
        "activity",
        "task",
        "step",
        "work_content",
    )

    # Setup / changeover time.
    mapping["setup_minutes"] = find_col(
        "setup",
        "setuptime",
        "changeover",
        "chgover",
        "段取り",
        "段取",
        "work_setup",
        "work_setuptime",
    )
    mapping["setup_start_time"] = find_col("work_operationsetupstarttime", "setupstarttime")

    return mapping


def _get_val(mapping: Mapping[str, str], row: Mapping[str, Any], key: str) -> str:
    col = mapping.get(key)
    if not col:
        return ""
    v = row.get(col, "")
    return str(v).strip() if v is not None else ""


def parse_schedule_upload_row(
    row: Mapping[str, Any], mapping: Mapping[str, str]
) -> Optional[Dict[str, Any]]:
    """
    Parse one uploaded CSV row into schedule table values.
    Returns None if the row is not processable (e.g. missing start_time).
    """
    start_raw = _get_val(mapping, row, "start_time")
    end_raw = _get_val(mapping, row, "end_time")
    start_dt = parse_datetime(start_raw)
    end_dt = parse_datetime(end_raw)
    if not start_dt:
        return None

    machine_id = _get_val(mapping, row, "machine_id") or _get_val(mapping, row, "machine_name")
    machine_name = (
        _get_val(mapping, row, "machine_name") or _get_val(mapping, row, "machine_id") or "Unknown"
    )

    qty_raw = _get_val(mapping, row, "quantity")
    qty: Optional[float]
    try:
        qty = float(qty_raw) if qty_raw else None
        if qty is not None and math.isnan(qty):
            qty = None
    except ValueError:
        qty = None

    # Setup time: either explicit duration column, or difference between setup start and production start.
    setup_raw = _get_val(mapping, row, "setup_minutes")
    setup_minutes = parse_duration_minutes(setup_raw)
    if setup_minutes is None:
        setup_start_raw = _get_val(mapping, row, "setup_start_time")
        setup_dt = parse_datetime(setup_start_raw)
        if setup_dt and start_dt and setup_dt < start_dt:
            setup_minutes = (start_dt - setup_dt).total_seconds() / 60.0

    return {
        "order_id": _get_val(mapping, row, "order_id"),
        "order_item_code": _get_val(mapping, row, "order_item_code"),
        "operation_id": _get_val(mapping, row, "operation_id"),
        "next_operation_id": _get_val(mapping, row, "next_operation_id"),
        "operation_code": _get_val(mapping, row, "operation_code"),
        "next_operation_code": _get_val(mapping, row, "next_operation_code"),
        "operation_out_item": _get_val(mapping, row, "operation_out_item"),
        "item_id": _get_val(mapping, row, "item_id"),
        "item_name": _get_val(mapping, row, "item_name"),
        "machine_id": machine_id,
        "machine_name": machine_name,
        "start_time": start_dt.strftime("%Y-%m-%d %H:%M:%S") if start_dt else None,
        "end_time": end_dt.strftime("%Y-%m-%d %H:%M:%S") if end_dt else None,
        "quantity": qty,
        "status": _get_val(mapping, row, "status") or "Scheduled",
        "process_name": _get_val(mapping, row, "process_name"),
        "setup_minutes": setup_minutes,
    }

