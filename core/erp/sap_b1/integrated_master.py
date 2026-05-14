"""SAP B1 (SQL Server) 用 Integrated Master 取得処理。

入力: ``dbo.ITT1_TMP`` + ``dbo.OITT_TMP``
出力: 7 列形式 (P_ITM_CD / PROCESS_NO / PROCESS_CD / INST_TYP / INST_CD /
              ITM_RESOURCE / PRODUCTION) — Bridge の integrated_master.csv と同じ。

正規化済みの dict 行を返すところまでを担当し、SQLite への永続化や
出力フォーマット整形は ``core.integrated_master`` 側のロジックに委譲する。
"""

from __future__ import annotations

from typing import Any, Dict, List, Tuple

from core.integrated_master import build_integrated_records, replace_integrates_in_sqlite

from .schema import bracket_ident, pick_column, sqlserver_table_columns


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
    itt1_cols = sqlserver_table_columns(pyodbc_conn, "ITT1_TMP")
    oitt_cols = sqlserver_table_columns(pyodbc_conn, "OITT_TMP")
    if not itt1_cols:
        raise RuntimeError("dbo.ITT1_TMP was not found, or column information could not be retrieved.")
    if not oitt_cols:
        raise RuntimeError("dbo.OITT_TMP was not found, or column information could not be retrieved.")

    parent = pick_column(itt1_cols, _ITT1_CANDIDATES["parent_item"])
    comp = pick_column(itt1_cols, _ITT1_CANDIDATES["component_code"])
    if not parent or not comp:
        raise RuntimeError(
            "No parent item or component item column was found in ITT1_TMP. "
            f"Actual columns: {', '.join(itt1_cols)}. "
            "Column names should be close to SAP export names, such as Parent Item or Component Code."
        )

    o_parent = pick_column(oitt_cols, _OITT_PARENT_CANDIDATES)
    if not o_parent:
        raise RuntimeError(
            "No parent item column was found in OITT_TMP. "
            f"Actual columns: {', '.join(oitt_cols)}."
        )

    q_parent = bracket_ident(parent)
    q_comp = bracket_ident(comp)
    q_o_parent = bracket_ident(o_parent)

    q_ln = pick_column(itt1_cols, _ITT1_CANDIDATES["line_no"])
    q_vo = pick_column(itt1_cols, _ITT1_CANDIDATES["visual_order"])
    q_ce = pick_column(itt1_cols, _ITT1_CANDIDATES["component_element_number"])
    q_qty = pick_column(itt1_cols, _ITT1_CANDIDATES["quantity"])
    q_wh = pick_column(itt1_cols, _ITT1_CANDIDATES["warehouse"])
    q_st = pick_column(itt1_cols, _ITT1_CANDIDATES["stage_id"])

    if q_ln:
        line_no_expr = f"t.{bracket_ident(q_ln)}"
    else:
        line_no_expr = (
            "ROW_NUMBER() OVER (PARTITION BY CAST(t."
            + q_parent
            + " AS NVARCHAR(256)) ORDER BY (SELECT NULL))"
        )

    qty_expr = f"t.{bracket_ident(q_qty)}" if q_qty else "CAST(NULL AS DECIMAL(18, 6))"
    wh_expr = f"t.{bracket_ident(q_wh)}" if q_wh else "CAST(NULL AS NVARCHAR(256))"
    st_expr = f"t.{bracket_ident(q_st)}" if q_st else "CAST(NULL AS NVARCHAR(256))"

    order_parts: List[str] = [f"t.{q_parent}"]
    if q_vo:
        order_parts.append(f"CASE WHEN t.{bracket_ident(q_vo)} IS NULL THEN 1 ELSE 0 END")
        order_parts.append(f"t.{bracket_ident(q_vo)}")
    if q_ce:
        order_parts.append(f"CASE WHEN t.{bracket_ident(q_ce)} IS NULL THEN 1 ELSE 0 END")
        order_parts.append(f"t.{bracket_ident(q_ce)}")
    if q_ln:
        order_parts.append(f"t.{bracket_ident(q_ln)}")
    order_sql = ",\n        ".join(order_parts)

    return f"""
    SELECT
        {line_no_expr} AS line_no,
        t.{q_parent} AS parent_item,
        t.{q_comp} AS component_code,
        {qty_expr} AS quantity,
        {"t." + bracket_ident(q_vo) if q_vo else "CAST(NULL AS INT)"} AS visual_order,
        {"t." + bracket_ident(q_ce) if q_ce else "CAST(NULL AS INT)"} AS component_element_number,
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
    """SQL Server から「親品目・子品目・数量・並び順キー」のフラット行を返す。

    列名は ``dbo.ITT1_TMP`` / ``dbo.OITT_TMP`` 実体に対して動的解決する
    (``Parent Item`` と ``ParentItem`` を等価扱いするなど)。
    """

    sql = _build_itt1_oitt_sql(pyodbc_conn)
    cur = pyodbc_conn.cursor()
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    out: List[Dict[str, Any]] = []
    for raw in cur.fetchall():
        out.append(dict(zip(cols, raw)))
    return out


def fetch_integrated_master_rows_from_sqlserver(
    pyodbc_conn,
) -> List[Dict[str, Any]]:
    """CSV 出力用の Integrated Master dict 行を返す (列名は ``INTEGRATED_HEADERS``)。"""

    flat = fetch_itt1_with_oitt_parents(pyodbc_conn)
    return build_integrated_records(flat)


def sync_from_sqlserver_to_sqlite(
    *,
    server: str,
    database: str,
    user: str,
    password: str,
    sqlite_db_path: str,
) -> int:
    """SQL Server から SAP B1 統合マスタを取得し、schedule.db ``integrates`` を置き換える。"""

    from .connection import connect_sqlserver

    pyodbc_conn = connect_sqlserver(server, database, user, password, timeout=30)
    try:
        flat = fetch_itt1_with_oitt_parents(pyodbc_conn)
        records = build_integrated_records(flat)
        return replace_integrates_in_sqlite(sqlite_db_path, records)
    finally:
        pyodbc_conn.close()
