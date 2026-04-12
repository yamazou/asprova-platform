"""
Build order_table.csv rows from dbo.SHIP_SCH (SAP B1).

Output columns match order_table.csv: REQ_NO, ITM_CD, DLV_DT, REQ_QTY, CUST_CD.
REQ_NO is DocumentNumber + '-' + line (# column), e.g. 5221-1.
DLV_DT is formatted like 3/18/2025 (M/d/yyyy). Column names are resolved from SHIP_SCH layout.
"""
from __future__ import annotations

from typing import Any, List, Tuple

from core.sap_integrated_master import (
    _bracket_ident,
    _pick_column,
    _sqlserver_table_columns,
)

_DOC_CANDIDATES: Tuple[str, ...] = (
    "Document Number",
    "DocumentNumber",
    "DocNum",
    "DocEntry",
    "Document",
    "REQ_NO",
)

_ITEM_CANDIDATES: Tuple[str, ...] = (
    "ItemCode",
    "Item Code",
    "Item No.",
    "ItemNo",
    "ITM_CD",
)

_SHIP_DATE_CANDIDATES: Tuple[str, ...] = (
    "ShipDate",
    "Ship Date",
    "Delivery Date",
    "DeliveryDate",
    "DLV_DT",
    "DueDate",
)

_QTY_CANDIDATES: Tuple[str, ...] = (
    "OpenQty",
    "Open Qty",
    "Open Quantity",
    "Quantity",
    "Qty",
    "REQ_QTY",
)

_CUST_CANDIDATES: Tuple[str, ...] = (
    "Customer/Vendor Code",
    "Customer/Vendor code",
    "Customer Code",
    "CustomerCode",
    "CardCode",
    "CUST_CD",
)

_LINE_CANDIDATES: Tuple[str, ...] = ("#", "No", "RowNo", "Line", "LineNo")


def fetch_order_rows_from_sqlserver(pyodbc_conn) -> List[Tuple[Any, Any, Any, Any, Any]]:
    """Return (REQ_NO, ITM_CD, DLV_DT, REQ_QTY, CUST_CD) rows."""
    cols = _sqlserver_table_columns(pyodbc_conn, "SHIP_SCH")
    if not cols:
        raise RuntimeError("dbo.SHIP_SCH が見つからないか、列情報を取得できません。")

    doc = _pick_column(cols, _DOC_CANDIDATES)
    itm = _pick_column(cols, _ITEM_CANDIDATES)
    sdt = _pick_column(cols, _SHIP_DATE_CANDIDATES)
    qty = _pick_column(cols, _QTY_CANDIDATES)
    if not doc or not itm or not sdt or not qty:
        raise RuntimeError(
            "SHIP_SCH に Document / Item / ShipDate / Quantity 系の列が見つかりません。"
            f" 実際の列: {', '.join(cols)}。"
        )

    cust = _pick_column(cols, _CUST_CANDIDATES)
    ln = _pick_column(cols, _LINE_CANDIDATES)
    qd = _bracket_ident(doc)
    qi = _bracket_ident(itm)
    qs = _bracket_ident(sdt)
    qq = _bracket_ident(qty)
    qc = _bracket_ident(cust) if cust else None

    doc_trim = f"LTRIM(RTRIM(CAST(t.{qd} AS NVARCHAR(256))))"
    item_trim = f"LTRIM(RTRIM(CAST(t.{qi} AS NVARCHAR(256))))"

    date_expr = (
        f"COALESCE("
        f"FORMAT(TRY_CAST(t.{qs} AS datetime2), N'M/d/yyyy'), "
        f"LTRIM(RTRIM(CAST(t.{qs} AS NVARCHAR(64)))), "
        f"N''"
        f")"
    )

    cust_expr = (
        f"LTRIM(RTRIM(COALESCE(CAST(t.{qc} AS NVARCHAR(256)), N'')))"
        if qc
        else "CAST(N'' AS NVARCHAR(256))"
    )

    if ln:
        qln = _bracket_ident(ln)
        ln_int_expr = f"TRY_CONVERT(INT, ROUND(TRY_CONVERT(FLOAT, t.{qln}), 0))"
    else:
        ln_int_expr = "CAST(NULL AS INT)"

    sql = f"""
    SELECT
        CONCAT(s._doc, N'-', CAST(COALESCE(s._ln_int, s._rn) AS NVARCHAR(32))) AS REQ_NO,
        s._item AS ITM_CD,
        s._dlv AS DLV_DT,
        s._qty AS REQ_QTY,
        s._cust AS CUST_CD
    FROM (
        SELECT
            {doc_trim} AS _doc,
            {item_trim} AS _item,
            {date_expr} AS _dlv,
            CAST(t.{qq} AS DECIMAL(18, 6)) AS _qty,
            {cust_expr} AS _cust,
            {ln_int_expr} AS _ln_int,
            ROW_NUMBER() OVER (
                PARTITION BY {doc_trim}
                ORDER BY
                    TRY_CAST(t.{qs} AS datetime2),
                    {item_trim},
                    {ln_int_expr}
            ) AS _rn
        FROM dbo.SHIP_SCH AS t
        WHERE {doc_trim} <> N''
          AND {item_trim} <> N''
    ) AS s
    ORDER BY s._doc, COALESCE(s._ln_int, s._rn), s._item
    """
    cur = pyodbc_conn.cursor()
    cur.execute(sql)
    return [tuple(row) for row in cur.fetchall()]
