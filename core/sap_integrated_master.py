"""
Build Asprova Integrated Master rows from SAP B1 staging tables ITT1_TMP + OITT_TMP.

Output matches the 7-column layout used by integrated_master.csv (same as Bridge):
  P_ITM_CD, PROCESS_NO, PROCESS_CD, INST_TYP, INST_CD, ITM_RESOURCE, PRODUCTION

- I rows: BOM lines from ITT1_TMP (inner join OITT_TMP on [Parent Item]).
  Quantities follow the reference CSV style (e.g. 1, .18, .018 — leading dot when |q| < 1).
- U rows: ITT1_TMP has only parent/child BOM data, so for each parent we emit one INST_TYP=U,
  INST_CD=M row with PROCESS_NO=10, PROCESS_CD=10, ITM_RESOURCE=PROD, PRODUCTION=1D (fixed).
  Set SAP_INTEGRATED_U_ROWS=0 to omit U rows.
- After each INST_TYP=I row, append one U/M row: P_ITM_CD = that I row's ITM_RESOURCE
  (the component code); PROCESS_NO/CD = 5; ITM_RESOURCE = SUPPLIER; PRODUCTION = 0D (fixed).
  Skipped when the I row's ITM_RESOURCE is empty.
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
    """
    Match integrated_master.csv style for INST_TYP I:
    integers as '1', fractions under 1 as '.18' / '.018' (no leading zero).
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
    # |q| < 1 → ".18" style
    if ad < 1:
        frac = format(ad, "f").rstrip("0").rstrip(".")
        if frac.startswith("0."):
            return sign + "." + frac[2:]
        return sign + frac
    s = format(ad, "f").rstrip("0").rstrip(".")
    return sign + s


def integrated_u_rows_enabled() -> bool:
    """Set SAP_INTEGRATED_U_ROWS=0 to emit only I rows (reference CSV often includes U rows)."""
    raw = (os.environ.get("SAP_INTEGRATED_U_ROWS") or "1").strip().lower()
    return raw not in ("0", "false", "no", "off")


def append_supplier_use_lines_after_inputs(
    records: Sequence[Mapping[str, Any]],
) -> List[Dict[str, Any]]:
    """
    For each INST_TYP=I row, append U/M with P_ITM_CD = that row's ITM_RESOURCE,
    PROCESS_NO=5, PROCESS_CD='5', INST_TYP='U', INST_CD='M',
    ITM_RESOURCE='SUPPLIER', PRODUCTION='0D'.
    """
    out: List[Dict[str, Any]] = []
    for r in records:
        out.append(dict(r))
        inst = str(r.get("INST_TYP") or "").strip().upper()
        if inst != "I":
            continue
        comp = r.get("ITM_RESOURCE")
        if comp is None or str(comp).strip() == "":
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


def _row_get(row: Mapping[str, Any], *keys: str) -> Any:
    for k in keys:
        if k in row and row[k] is not None:
            return row[k]
    return None


def _norm_col_key(name: str) -> str:
    """Lowercase alphanumerics only — matches 'Parent Item', 'ParentItem', 'parent_item'."""
    return "".join(ch.lower() for ch in name if ch.isalnum())


def _bracket_ident(col: str) -> str:
    return "[" + col.replace("]", "]]") + "]"


def _sqlserver_table_columns(pyodbc_conn, table: str, schema: str = "dbo") -> List[str]:
    cur = pyodbc_conn.cursor()
    cur.execute(
        """
        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """,
        (schema, table),
    )
    return [row[0] for row in cur.fetchall()]


def _pick_column(available: Sequence[str], candidates: Sequence[str]) -> Optional[str]:
    by_norm: Dict[str, str] = {}
    for c in available:
        n = _norm_col_key(c)
        if n not in by_norm:
            by_norm[n] = c
    for cand in candidates:
        n = _norm_col_key(cand)
        if n in by_norm:
            return by_norm[n]
    return None


# Preference order per logical field (SAP B1 Excel export uses spaced names).
_ITT1_CANDIDATES: Dict[str, Tuple[str, ...]] = {
    "line_no": ("#", "No", "RowNo", "Line", "LineNo", "LinNum"),
    "parent_item": ("Parent Item", "ParentItem", "Parent_Item", "Father", "ParentCode"),
    "component_code": (
        "Component Code",
        "ComponentCode",
        "Component_Item",
        "ChildItem",
        "ItemCode",
    ),
    "quantity": ("Quantity", "Qty", "QTY"),
    "visual_order": ("Visual Order", "VisualOrder", "VisOrder"),
    "component_element_number": (
        "Component Element Number",
        "ComponentElementNumber",
        "CompElemNum",
    ),
    "warehouse": ("Warehouse", "WhsCode", "Whse", "WarehouseCode"),
    "stage_id": ("Stage ID", "StageID", "Stage_Id"),
}

_OITT_PARENT_CANDIDATES: Tuple[str, ...] = (
    "Parent Item",
    "ParentItem",
    "Parent_Item",
    "Code",
    "ItemCode",
    "Father",
    "ParentCode",
)


def _build_itt1_oitt_sql(pyodbc_conn) -> str:
    itt1_cols = _sqlserver_table_columns(pyodbc_conn, "ITT1_TMP")
    oitt_cols = _sqlserver_table_columns(pyodbc_conn, "OITT_TMP")
    if not itt1_cols:
        raise RuntimeError("dbo.ITT1_TMP が見つからないか、列情報を取得できません。")
    if not oitt_cols:
        raise RuntimeError("dbo.OITT_TMP が見つからないか、列情報を取得できません。")

    parent = _pick_column(itt1_cols, _ITT1_CANDIDATES["parent_item"])
    comp = _pick_column(itt1_cols, _ITT1_CANDIDATES["component_code"])
    if not parent or not comp:
        raise RuntimeError(
            "ITT1_TMP に親品目・子品目列が見つかりません。"
            f" 実際の列: {', '.join(itt1_cols)}。"
            " SAP エクスポート列名（例: Parent Item, Component Code）に近い名前が必要です。"
        )

    o_parent = _pick_column(oitt_cols, _OITT_PARENT_CANDIDATES)
    if not o_parent:
        raise RuntimeError(
            "OITT_TMP に親品目列が見つかりません。"
            f" 実際の列: {', '.join(oitt_cols)}。"
        )

    q_parent = _bracket_ident(parent)
    q_comp = _bracket_ident(comp)
    q_o_parent = _bracket_ident(o_parent)

    q_ln = _pick_column(itt1_cols, _ITT1_CANDIDATES["line_no"])
    q_vo = _pick_column(itt1_cols, _ITT1_CANDIDATES["visual_order"])
    q_ce = _pick_column(itt1_cols, _ITT1_CANDIDATES["component_element_number"])
    q_qty = _pick_column(itt1_cols, _ITT1_CANDIDATES["quantity"])
    q_wh = _pick_column(itt1_cols, _ITT1_CANDIDATES["warehouse"])
    q_st = _pick_column(itt1_cols, _ITT1_CANDIDATES["stage_id"])

    if q_ln:
        line_no_expr = f"t.{_bracket_ident(q_ln)}"
    else:
        line_no_expr = (
            "ROW_NUMBER() OVER (PARTITION BY CAST(t."
            + q_parent
            + " AS NVARCHAR(256)) ORDER BY (SELECT NULL))"
        )

    qty_expr = f"t.{_bracket_ident(q_qty)}" if q_qty else "CAST(NULL AS DECIMAL(18, 6))"
    wh_expr = f"t.{_bracket_ident(q_wh)}" if q_wh else "CAST(NULL AS NVARCHAR(256))"
    st_expr = f"t.{_bracket_ident(q_st)}" if q_st else "CAST(NULL AS NVARCHAR(256))"

    order_parts: List[str] = [f"t.{q_parent}"]
    if q_vo:
        order_parts.append(f"CASE WHEN t.{_bracket_ident(q_vo)} IS NULL THEN 1 ELSE 0 END")
        order_parts.append(f"t.{_bracket_ident(q_vo)}")
    if q_ce:
        order_parts.append(
            f"CASE WHEN t.{_bracket_ident(q_ce)} IS NULL THEN 1 ELSE 0 END"
        )
        order_parts.append(f"t.{_bracket_ident(q_ce)}")
    if q_ln:
        order_parts.append(f"t.{_bracket_ident(q_ln)}")
    order_sql = ",\n        ".join(order_parts)

    return f"""
    SELECT
        {line_no_expr} AS line_no,
        t.{q_parent} AS parent_item,
        t.{q_comp} AS component_code,
        {qty_expr} AS quantity,
        {"t." + _bracket_ident(q_vo) if q_vo else "CAST(NULL AS INT)"} AS visual_order,
        {"t." + _bracket_ident(q_ce) if q_ce else "CAST(NULL AS INT)"} AS component_element_number,
        {wh_expr} AS warehouse,
        {st_expr} AS stage_id
    FROM dbo.ITT1_TMP AS t
    INNER JOIN dbo.OITT_TMP AS o
      ON CAST(o.{q_o_parent} AS NVARCHAR(256)) = CAST(t.{q_parent} AS NVARCHAR(256))
    WHERE LTRIM(RTRIM(COALESCE(CAST(t.{q_comp} AS NVARCHAR(256)), N''))) <> N''
    ORDER BY
        {order_sql}
    """


def fetch_itt1_with_oitt_parents(pyodbc_conn) -> List[Dict[str, Any]]:
    """
    Return flat rows from SQL Server: parent item, component, qty, sort keys.
    Resolves column names against dbo.ITT1_TMP / dbo.OITT_TMP (handles ParentItem vs Parent Item).
    """
    sql = _build_itt1_oitt_sql(pyodbc_conn)
    cur = pyodbc_conn.cursor()
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    out: List[Dict[str, Any]] = []
    for raw in cur.fetchall():
        out.append(dict(zip(cols, raw)))
    return out


def build_integrated_records(flat_rows: Sequence[Mapping[str, Any]]) -> List[Dict[str, Any]]:
    """
    Group by parent_item; assign In1..InN in sort order within each parent, then U rows.
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


def replace_integrates_in_sqlite(db_path: str, records: Sequence[Mapping[str, Any]]) -> int:
    """
    DELETE all rows in integrates, then INSERT records. Returns row count inserted.
    """
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


def sync_from_sqlserver_to_sqlite(
    *,
    server: str,
    database: str,
    user: str,
    password: str,
    sqlite_db_path: str,
) -> int:
    """Connect to SQL Server, build Integrated Master rows, write to schedule.db integrates."""
    from core.sqlserver_conn import connect_sqlserver

    pyodbc_conn = connect_sqlserver(server, database, user, password, timeout=30)
    try:
        flat = fetch_itt1_with_oitt_parents(pyodbc_conn)
        records = build_integrated_records(flat)
        return replace_integrates_in_sqlite(sqlite_db_path, records)
    finally:
        pyodbc_conn.close()


def fetch_integrated_master_rows_from_sqlserver(pyodbc_conn) -> List[Dict[str, Any]]:
    """Return Integrated Master dict rows (same keys as INTEGRATED_HEADERS) for CSV export."""
    flat = fetch_itt1_with_oitt_parents(pyodbc_conn)
    return build_integrated_records(flat)
