"""ERP 非依存の Integrated Master (BOM + 工程行) ドメインロジック。

ここに置くもの:
    - Asprova の Integrated Master CSV / SQLite ``integrates`` テーブルの形式定義
    - flat な BOM 行 → Integrated Master 行への正規化処理
    - SQLite 側の DDL / replace 処理

ERP からのデータ取得は ``core.erp.<system>.integrated_master`` 側で行い、
取得後の正規化・永続化はこのモジュールを使うことで、
新しい ERP を増やしても出力フォーマットは 1 箇所に集約できる。
"""

from __future__ import annotations

import math
import os
import sqlite3
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple


INTEGRATED_HEADERS: Tuple[str, ...] = (
    "P_ITM_CD",
    "PROCESS_NO",
    "PROCESS_CD",
    "INST_TYP",
    "INST_CD",
    "ITM_RESOURCE",
    "PRODUCTION",
)


def _to_decimal(value: Any) -> Optional[Decimal]:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, float):
        if math.isnan(value):
            return None
        return Decimal(str(value))
    if isinstance(value, int):
        return Decimal(value)
    s = str(value).strip().replace(",", "")
    if s == "":
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def format_integrated_i_production(value: Any) -> str:
    """Match integrated_master.csv style for INST_TYP I.

    integers as ``'1'``, fractions under 1 as ``'.18'`` / ``'.018'``
    (no leading zero).
    """

    d = _to_decimal(value)
    if d is None:
        return ""
    if d.is_zero():
        return "0"
    sign = "-" if d < 0 else ""
    ad = abs(d)
    try:
        if (ad % 1) == 0:
            return sign + str(int(ad))
    except (InvalidOperation, ValueError, TypeError):
        pass
    if ad < 1:
        frac = format(ad, "f").rstrip("0").rstrip(".")
        if frac.startswith("0."):
            return sign + "." + frac[2:]
        return sign + frac
    s = format(ad, "f").rstrip("0").rstrip(".")
    return sign + s


def integrated_u_rows_enabled() -> bool:
    """Set ``SAP_INTEGRATED_U_ROWS=0`` to emit only I rows."""

    raw = (os.environ.get("SAP_INTEGRATED_U_ROWS") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def _row_get(row: Mapping[str, Any], *keys: str) -> Any:
    for k in keys:
        if k in row and row[k] is not None:
            return row[k]
    return None


def append_supplier_use_lines_after_inputs(
    records: Sequence[Mapping[str, Any]],
) -> List[Dict[str, Any]]:
    """For each ``INST_TYP=I`` row, append a ``U/M`` row whose ``P_ITM_CD``
    is that row's ``ITM_RESOURCE``, only when that ``ITM_RESOURCE`` does not
    already appear as a parent.
    """

    out: List[Dict[str, Any]] = []
    existing_parent_items = {
        str(r.get("P_ITM_CD") or "").strip()
        for r in records
        if str(r.get("P_ITM_CD") or "").strip() != ""
    }
    for r in records:
        out.append(dict(r))
        inst = str(r.get("INST_TYP") or "").strip().upper()
        if inst != "I":
            continue
        comp = r.get("ITM_RESOURCE")
        if comp is None or str(comp).strip() == "":
            continue
        if str(comp).strip() in existing_parent_items:
            continue
        out.append(
            {
                "P_ITM_CD": str(comp).strip(),
                "PROCESS_NO": 5,
                "PROCESS_CD": "5",
                "INST_TYP": "U",
                "INST_CD": "M",
                "ITM_RESOURCE": "SUPPLIER",
                "PRODUCTION": "0D",
            }
        )
    return out


def build_integrated_records(
    flat_rows: Sequence[Mapping[str, Any]],
) -> List[Dict[str, Any]]:
    """Group by ``parent_item``; assign ``In1..InN`` in sort order within each
    parent, then optionally append U rows. Finally append SUPPLIER/USE lines.

    入力は ERP 側が dict 化した「親品目・子品目・数量・並び順キー」のフラット行。
    """

    by_parent: Dict[str, List[Mapping[str, Any]]] = defaultdict(list)
    for r in flat_rows:
        p = _row_get(r, "parent_item", "Parent Item")
        if p is None or str(p).strip() == "":
            continue
        by_parent[str(p).strip()].append(r)

    records: List[Dict[str, Any]] = []
    for parent in sorted(by_parent.keys()):
        items = by_parent[parent]

        def sort_key(m: Mapping[str, Any]) -> Tuple:
            vo = _row_get(m, "visual_order", "Visual Order")
            ce = _row_get(m, "component_element_number", "Component Element Number")
            ln = _row_get(m, "line_no", "#")
            vo_n = vo if vo is not None else 10**9
            ce_n = ce if ce is not None else 10**9
            ln_n = ln if ln is not None else 10**9
            try:
                vo_n = int(vo_n)
            except (TypeError, ValueError):
                vo_n = 10**9
            try:
                ce_n = int(ce_n)
            except (TypeError, ValueError):
                ce_n = 10**9
            try:
                ln_n = int(ln_n)
            except (TypeError, ValueError):
                ln_n = 10**9
            return (vo_n, ce_n, ln_n)

        items.sort(key=sort_key)
        for i, m in enumerate(items, start=1):
            comp = _row_get(m, "component_code", "Component Code")
            qty = _row_get(m, "quantity", "Quantity")
            records.append(
                {
                    "P_ITM_CD": parent,
                    "PROCESS_NO": 10,
                    "PROCESS_CD": "10",
                    "INST_TYP": "I",
                    "INST_CD": f"In{i}",
                    "ITM_RESOURCE": str(comp).strip(),
                    "PRODUCTION": format_integrated_i_production(qty),
                }
            )
        if integrated_u_rows_enabled():
            records.append(
                {
                    "P_ITM_CD": parent,
                    "PROCESS_NO": 10,
                    "PROCESS_CD": "10",
                    "INST_TYP": "U",
                    "INST_CD": "M",
                    "ITM_RESOURCE": "PROD",
                    "PRODUCTION": "1D",
                }
            )
    return append_supplier_use_lines_after_inputs(records)


def ensure_integrates_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS integrates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            P_ITM_CD TEXT NOT NULL,
            PROCESS_NO INTEGER,
            PROCESS_CD TEXT,
            INST_TYP TEXT,
            INST_CD TEXT,
            ITM_RESOURCE TEXT,
            PRODUCTION TEXT,
            synced_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_integrates_p_itm_cd ON integrates (P_ITM_CD)"
    )


def replace_integrates_in_sqlite(
    db_path: str, records: Sequence[Mapping[str, Any]]
) -> int:
    """``integrates`` を全削除→記録を ``INSERT``。挿入件数を返す。"""

    conn = sqlite3.connect(db_path)
    try:
        ensure_integrates_table(conn)
        conn.execute("DELETE FROM integrates")
        for r in records:
            conn.execute(
                """
                INSERT INTO integrates
                (P_ITM_CD, PROCESS_NO, PROCESS_CD, INST_TYP, INST_CD, ITM_RESOURCE, PRODUCTION)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    r["P_ITM_CD"],
                    r["PROCESS_NO"],
                    r["PROCESS_CD"],
                    r["INST_TYP"],
                    r["INST_CD"],
                    r["ITM_RESOURCE"],
                    r["PRODUCTION"],
                ),
            )
        conn.commit()
        return len(records)
    finally:
        conn.close()
