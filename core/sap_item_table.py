"""
Build Item Table CSV rows (ITM_CD, ITM_NM, MAX_LOT_UNIT_QTY) from dbo.OITM_TMP (SAP B1).

MAX_LOT_UNIT_QTY is always NULL (blank in CSV). Column names are resolved like other SAP staging loaders.
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
    "Code",
)

_ITEM_NM_CANDIDATES: Tuple[str, ...] = (
    "Item Description",
    "ItemDescription",
    "Description",
    "Item Name",
    "ItemName",
    "ItemNm",
    "Foreign Name",
    "ForeignName",
)


def fetch_item_table_rows_from_sqlserver(pyodbc_conn) -> List[Tuple[Any, Any, Any]]:
    """Return (ITM_CD, ITM_NM, MAX_LOT_UNIT_QTY) rows; third column is always None."""
    cols = _sqlserver_table_columns(pyodbc_conn, "OITM_TMP")
    if not cols:
        raise RuntimeError("dbo.OITM_TMP が見つからないか、列情報を取得できません。")

    itm_cd = _pick_column(cols, _ITEM_CD_CANDIDATES)
    if not itm_cd:
        raise RuntimeError(
            "OITM_TMP に品目コード列が見つかりません。"
            f" 実際の列: {', '.join(cols)}。"
        )

    itm_nm = _pick_column(cols, _ITEM_NM_CANDIDATES)
    q_cd = _bracket_ident(itm_cd)
    if itm_nm:
        q_nm = _bracket_ident(itm_nm)
        nm_sql = f"LTRIM(RTRIM(COALESCE(CAST(t.{q_nm} AS NVARCHAR(MAX)), N'')))"
    else:
        nm_sql = "CAST(N'' AS NVARCHAR(MAX))"

    sql = f"""
    SELECT DISTINCT
        LTRIM(RTRIM(CAST(t.{q_cd} AS NVARCHAR(256)))) AS ITM_CD,
        {nm_sql} AS ITM_NM,
        CAST(NULL AS NVARCHAR(64)) AS MAX_LOT_UNIT_QTY
    FROM dbo.OITM_TMP AS t
    WHERE LTRIM(RTRIM(COALESCE(CAST(t.{q_cd} AS NVARCHAR(256)), N''))) <> N''
    ORDER BY ITM_CD
    """
    cur = pyodbc_conn.cursor()
    cur.execute(sql)
    return [tuple(row) for row in cur.fetchall()]
