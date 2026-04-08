from __future__ import annotations

import math
from datetime import datetime
from typing import Any, Dict, List, Mapping, Optional, Tuple


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
    # work_type before work_code so exports with both Work_Code + Work_Type round-trip correctly.
    mapping["process_name"] = find_col(
        "work_type",
        "work_operationoutmainitemcode",
        "work_code",
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

    # Actuals (optional CSV columns)
    mapping["actual_start"] = find_exact("actual_start") or find_col(
        "actual_start",
        "work_actualstarttime",
        "actualstarttime",
        "result_actualstart",
        "actual_start_time",
    )
    mapping["actual_end"] = find_exact("actual_end") or find_col(
        "actual_end",
        "work_actualendtime",
        "actualendtime",
        "result_actualend",
        "actual_end_time",
    )
    mapping["actual_resource"] = find_exact("actual_resource") or find_col(
        "actual_resource",
        "work_actualmainrescode",
        "actualmainrescode",
        "result_actualres",
        "actual_res",
    )

    # Resource / skill grouping (optional CSV columns)
    # WorkUser_Group (Asprova export) before plain "Group" so both can coexist in one CSV.
    mapping["work_group"] = (
        find_exact("workuser_group")
        or find_col("workuser_group", "workuser.group", "workuser group")
        or find_exact("group")
        or find_col(
            "work_group",
            "resource_group",
            "op_group",
            "グループ",
        )
    )
    mapping["work_user_res_order"] = find_exact("workuser_resorder") or find_col(
        "workuser_resorder",
        "workuser_res_order",
    )
    mapping["delivery_date"] = find_exact("workuser_deliverydate") or find_col(
        "workuser_deliverydate",
        "workuser_delivery_date",
    )
    mapping["delivery_order_no"] = find_exact("workuser_deliveryorderno") or find_col(
        "workuser_deliveryorderno",
        "workuser_delivery_order_no",
    )
    mapping["delivery_item"] = find_exact("workuser_deliveryitem") or find_col(
        "workuser_deliveryitem",
        "workuser_delivery_item",
    )
    mapping["delivery_item_name"] = find_exact("workuser_deliveryitemname") or find_col(
        "workuser_deliveryitemname",
        "workuser_delivery_item_name",
    )
    mapping["min_skill"] = find_exact("min skill") or find_col(
        "min_skill",
        "minskill",
        "minimum skill",
        "最低スキル",
    )
    mapping["qc_skill"] = find_exact("qc skill") or find_col(
        "qc_skill",
        "qcskill",
        "q_c skill",
        "qc スキル",
    )

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

    actual_start_raw = _get_val(mapping, row, "actual_start")
    actual_end_raw = _get_val(mapping, row, "actual_end")
    actual_start_dt = parse_datetime(actual_start_raw)
    actual_end_dt = parse_datetime(actual_end_raw)
    actual_resource = _get_val(mapping, row, "actual_resource")

    wg = _get_val(mapping, row, "work_group")
    wro = _get_val(mapping, row, "work_user_res_order")
    dd = _get_val(mapping, row, "delivery_date")
    don = _get_val(mapping, row, "delivery_order_no")
    di = _get_val(mapping, row, "delivery_item")
    din = _get_val(mapping, row, "delivery_item_name")
    ms = _get_val(mapping, row, "min_skill")
    qs = _get_val(mapping, row, "qc_skill")

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
        "actual_start": (
            actual_start_dt.strftime("%Y-%m-%d %H:%M:%S") if actual_start_dt else None
        ),
        "actual_end": (
            actual_end_dt.strftime("%Y-%m-%d %H:%M:%S") if actual_end_dt else None
        ),
        "actual_resource": actual_resource or None,
        "work_group": wg or None,
        "work_user_res_order": wro or None,
        "delivery_date": dd or None,
        "delivery_order_no": don or None,
        "delivery_item": di or None,
        "delivery_item_name": din or None,
        "min_skill": ms or None,
        "qc_skill": qs or None,
    }


def _row_get(row: Any, key: str) -> Any:
    try:
        return row[key]
    except (KeyError, TypeError, IndexError):
        return None


def _format_csv_cell(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if math.isfinite(v) and abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return str(v)
    return str(v).strip()


# Result export: Work_Code + actual columns only.
RESULT_CSV_HEADERS: Tuple[str, ...] = (
    "Work_Code",
    "Actual_Start",
    "Actual_End",
    "Actual_Resource",
    "actual_quantity",
)


def result_csv_export_headers() -> List[str]:
    return list(RESULT_CSV_HEADERS)


def _format_result_export_datetime(v: Any) -> str:
    """Result CSV: Actual_Start / Actual_End as YYYY/MM/DD HH:MM:SS (e.g. 2026/03/05 03:12:00)."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    dt = parse_datetime(s)
    if not dt:
        return s
    return dt.strftime("%Y/%m/%d %H:%M:%S")


def schedule_row_to_result_csv_cells(row: Any) -> List[str]:
    oc = _row_get(row, "operation_code")
    pn = _row_get(row, "process_name")
    if oc is not None and str(oc).strip() != "":
        wc = oc
    else:
        wc = pn
    return [
        _format_csv_cell(wc),
        _format_result_export_datetime(_row_get(row, "actual_start")),
        _format_result_export_datetime(_row_get(row, "actual_end")),
        _format_csv_cell(_row_get(row, "actual_resource")),
        _format_csv_cell(_row_get(row, "actual_quantity")),
    ]

