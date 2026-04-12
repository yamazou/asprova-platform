"""
Build inventory_table.csv rows (INV_CD, ITM_CD, STK_QTY, INV_DT) from dbo.BEG_INV (SAP B1).

INV_CD is INV00001-style (FORMAT row number). INV_DT is always NULL (blank in CSV).
Column names are resolved like other SAP staging loaders.
"""
from __future__ import annotations

from typing import Any, List, Tuple

from core.sap_integrated_master import (
    _bracket_ident,
    _pick_column,
    _sqlserver_table_columns,
)

_ITEM_CD_CANDIDATES: Tuple[str, ...] = (
    "Item No.",
    "Item No",
    "ItemNo",
    "Item_Code",
    "Item Code",
    "ItemCode",
    "ITM_CD",
)

_STOCK_CANDIDATES: Tuple[str, ...] = (
    "In Stock",
    "InStock",
    "Stock",
    "STK_QTY",
    "On Hand",
    "OnHand",
    "Qty",
    "Quantity",
)

_WH_CANDIDATES: Tuple[str, ...] = (
    "Warehouse Code",
    "WarehouseCode",
    "WhsCode",
    "Whse",
)

_LINE_CANDIDATES: Tuple[str, ...] = ("#", "No", "RowNo", "Line", "LineNo")


def fetch_inventory_rows_from_sqlserver(pyodbc_conn) -> List[Tuple[Any, Any, Any, Any]]:
    """Return (INV_CD, ITM_CD, STK_QTY, INV_DT) rows; INV_DT is always None."""
    cols = _sqlserver_table_columns(pyodbc_conn, "BEG_INV")
    if not cols:
        raise RuntimeError("dbo.BEG_INV が見つからないか、列情報を取得できません。")

    itm = _pick_column(cols, _ITEM_CD_CANDIDATES)
    stk = _pick_column(cols, _STOCK_CANDIDATES)
    if not itm or not stk:
        raise RuntimeError(
            "BEG_INV に品目コード列・在庫数量列が見つかりません。"
            f" 実際の列: {', '.join(cols)}。"
        )

    qi = _bracket_ident(itm)
    qs = _bracket_ident(stk)
    item_expr = f"LTRIM(RTRIM(CAST(t.{qi} AS NVARCHAR(256))))"

    order_parts: List[str] = [item_expr]
    wh = _pick_column(cols, _WH_CANDIDATES)
    if wh:
        qw = _bracket_ident(wh)
        order_parts.append(f"LTRIM(RTRIM(CAST(t.{qw} AS NVARCHAR(256))))")
    ln = _pick_column(cols, _LINE_CANDIDATES)
    if ln:
        order_parts.append(f"t.{_bracket_ident(ln)}")
    ob = ", ".join(order_parts)

    sql = f"""
    SELECT
        CONCAT(
            'INV',
            FORMAT(ROW_NUMBER() OVER (ORDER BY {ob}), '00000')
        ) AS INV_CD,
        {item_expr} AS ITM_CD,
        CAST(t.{qs} AS DECIMAL(18, 6)) AS STK_QTY,
        CAST(NULL AS NVARCHAR(16)) AS INV_DT
    FROM dbo.BEG_INV AS t
    WHERE LTRIM(RTRIM(COALESCE(CAST(t.{qi} AS NVARCHAR(256)), N''))) <> N''
    ORDER BY {ob}
    """
    cur = pyodbc_conn.cursor()
    cur.execute(sql)
    return [tuple(row) for row in cur.fetchall()]
