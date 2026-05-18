"""Inventory Table 行を ITM_CD 単位に集計する。"""

from __future__ import annotations

import re
from collections import OrderedDict
from decimal import Decimal, InvalidOperation
from typing import Any, Sequence

_QTY_SKIP_TEXT = frozenset(
    {"on hand", "onhand", "stock", "-", "n/a", "na", ""}
)


def _normalize_itm_cd(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value).strip()
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        return str(int(float(s)))
    if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_.]*", s):
        return s.upper()
    return s


def _parse_stk_qty(value: Any) -> Decimal:
    if value is None:
        return Decimal(0)
    if isinstance(value, bool):
        return Decimal(0)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    s = str(value).strip()
    if not s or s.lower() in _QTY_SKIP_TEXT:
        return Decimal(0)
    s = s.replace(",", "")
    m = re.match(r"^([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)", s)
    if m:
        try:
            return Decimal(m.group(1))
        except InvalidOperation:
            return Decimal(0)
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal(0)


def _format_stk_qty(total: Decimal) -> str:
    if total.is_zero():
        return "0"
    if total == total.to_integral_value():
        return str(int(total))
    s = format(total.normalize(), "f")
    return s.rstrip("0").rstrip(".") or "0"


def _inv_cd_prefix(rows: Sequence[tuple]) -> str:
    for row in rows:
        if not row:
            continue
        code = str(row[0] or "").strip().upper()
        if code.startswith("WIP"):
            return "WIP"
    return "INV"


def aggregate_inventory_rows_by_itm_cd(
    rows: Sequence[tuple],
    *,
    inv_cd_prefix: str | None = None,
) -> list[tuple]:
    """同一 ``ITM_CD`` の行を 1 行にまとめ、``STK_QTY`` を合算する。

    出力列: ``(INV_CD, ITM_CD, STK_QTY, INV_DT)``。
    ``INV_CD`` は集計後に ``INV00001`` 形式で採番し直す（WIP 由来なら ``WIP`` 接頭辞）。
    """

    prefix = (inv_cd_prefix or _inv_cd_prefix(rows)).strip().upper() or "INV"
    grouped: OrderedDict[str, dict[str, Any]] = OrderedDict()

    for row in rows:
        if len(row) < 3:
            continue
        itm_cd = _normalize_itm_cd(row[1])
        if not itm_cd:
            continue
        if itm_cd not in grouped:
            grouped[itm_cd] = {"qty": Decimal(0), "inv_dt": ""}
        bucket = grouped[itm_cd]
        bucket["qty"] += _parse_stk_qty(row[2])
        inv_dt = str(row[3] or "").strip() if len(row) > 3 else ""
        if inv_dt and not bucket["inv_dt"]:
            bucket["inv_dt"] = inv_dt

    out: list[tuple] = []
    for seq, (itm_cd, bucket) in enumerate(grouped.items(), start=1):
        inv_cd = f"{prefix}{seq:05d}"
        out.append(
            (
                inv_cd,
                itm_cd,
                _format_stk_qty(bucket["qty"]),
                bucket["inv_dt"],
            )
        )
    return out
