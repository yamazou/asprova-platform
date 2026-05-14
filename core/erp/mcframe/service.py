"""mcframe (Oracle) 用 Bridge サービス実装。

Bridge アプリは ``apps/bridge/app.py`` から ``McframeBridgeService`` を生成し、
このオブジェクト経由で各種マスタ/トランザクションを取得する。
SQL 文はこのモジュール内に閉じ込めており、SQL Server / Excel 系のコードを
含めずに mcframe 顧客向け納品が可能。
"""

from __future__ import annotations

from typing import Any, Iterator

from core.integrated_master import (
    INTEGRATED_HEADERS,
    append_supplier_use_lines_after_inputs,
)
from core.erp._base import BridgeErpService


# ---------------------------------------------------------------------------
# Public column headers (ERP 横断で同じ列構成にする)
# ---------------------------------------------------------------------------


ITEM_TABLE_HEADERS = [
    "ITM_CD",
    "ITM_NM",
    "ITM_TYP",
    "MAX_LOT_UNIT_QTY",
]

ORDER_TABLE_HEADERS = [
    "REQ_NO",
    "ITM_CD",
    "DLV_DT",
    "REQ_QTY",
    "CUST_CD",
]

RESOURCE_TABLE_HEADERS = [
    "LINE_CD",
    "LINE_NM",
    "RESOURCE_GRP",
    "Sort_Order",
]

INVENTORY_TABLE_HEADERS = [
    "INV_CD",
    "ITM_CD",
    "STK_QTY",
    "INV_DT",
]


# ---------------------------------------------------------------------------
# SQL templates (PL/SQL は使わず、{schema} と {co_cd} を str.format で展開)
# ---------------------------------------------------------------------------


_MASTER_SQL = """
SELECT
    b.P_ITM_CD AS P_ITM_CD,
    10 AS PROCESS_NO,
    '10' AS PROCESS_CD,
    'I' AS INST_TYP,
    'In' || ROW_NUMBER() OVER (
      PARTITION BY b.P_ITM_CD
      ORDER BY b.C_ITM_CD
    ) AS INST_CD,
    b.C_ITM_CD AS ITM_RESOURCE,
    TO_CHAR(b.C_REQ_QTY) AS PRODUCTION
FROM
    {schema}.SM_BOM_ALL b
WHERE
    b.BOM_PTN = 1

UNION ALL

SELECT
    hl.ITM_CD AS P_ITM_CD,
    10 AS PROCESS_NO,
    '10' AS PROCESS_CD,
    'U' AS INST_TYP,
    'M' AS INST_CD,
    hl.LINE_CD AS ITM_RESOURCE,
    TO_CHAR(hl.CYCLE_TIME) || 'mp' AS PRODUCTION
FROM
    {schema}.SM_HINLINE_ALL hl
WHERE
    hl.CO_CD = '{co_cd}'
"""


_ITEM_TABLE_SQL = """
SELECT
    c.ITM_CD,
    c.ITM_NM,
    CASE TRIM(TO_CHAR(c.ITM_TYP))
        WHEN '1' THEN 'P'
        WHEN '2' THEN 'I'
        WHEN '5' THEN 'M'
        WHEN '6' THEN 'H'
        WHEN '7' THEN 'U'
        ELSE TRIM(TO_CHAR(c.ITM_TYP))
    END AS ITM_TYP,
    MAX(m.MAX_LOT_UNIT_QTY) AS MAX_LOT_UNIT_QTY
FROM
    {schema}.CM_HINMO_ALL c
    JOIN {schema}.SM_HINMOS_ALL m
      ON m.ITM_CD = c.ITM_CD
WHERE
    c.CO_CD = '{co_cd}'
    AND m.CO_CD = '{co_cd}'
GROUP BY
    c.ITM_CD,
    c.ITM_NM,
    c.ITM_TYP
ORDER BY
    c.ITM_CD
"""


_ORDER_TABLE_SQL = """
SELECT
    REQ_NO,
    ITM_CD,
    DLV_DT,
    REQ_QTY,
    CAST(NULL AS VARCHAR2(64)) AS CUST_CD
FROM
    {schema}.ST_SHOYO_ALL
ORDER BY
    REQ_NO,
    ITM_CD,
    DLV_DT
"""


_RESOURCE_TABLE_SQL = """
SELECT
    LINE_CD,
    LINE_NM,
    CASE
        WHEN LINE_CD LIKE 'M%' THEN 'INJECTION'
        ELSE ''
    END AS RESOURCE_GRP
FROM
    {schema}.SM_LINE_ALL
ORDER BY
    LINE_CD
"""


_INVENTORY_TABLE_SQL = """
SELECT
    'INV' || LPAD(ROW_NUMBER() OVER (ORDER BY ITM_CD), 5, '0') AS INV_CD,
    ITM_CD,
    STK_QTY,
    CAST(NULL AS VARCHAR2(10)) AS INV_DT
FROM
    {schema}.ST_GNZAIKO_ALL
ORDER BY
    ITM_CD
"""


# ---------------------------------------------------------------------------
# Service implementation
# ---------------------------------------------------------------------------


class McframeBridgeService(BridgeErpService):
    """mcframe (Oracle) 用 Bridge サービス。

    Args:
        oracle_user / oracle_password / oracle_dsn: oracledb 接続パラメータ
        oracle_schema: SM_BOM_ALL などの参照スキーマ
        mcframe_co_cd: mcframe 会社コード ({co_cd} へ展開)
    """

    def __init__(
        self,
        *,
        oracle_user: str,
        oracle_password: str,
        oracle_dsn: str,
        oracle_schema: str,
        mcframe_co_cd: str,
    ) -> None:
        if not (oracle_user and oracle_password):
            raise RuntimeError(
                "Not connected. Please connect to the database from Connect first."
            )
        if not oracle_schema:
            raise RuntimeError(
                "Not connected. Please connect to the database from Connect first."
            )
        self._user = oracle_user
        self._password = oracle_password
        self._dsn = oracle_dsn or "orcl"
        self._schema = oracle_schema
        self._co_cd = mcframe_co_cd or "J0001"

    # -- internal helpers -----------------------------------------------------

    @staticmethod
    def _ensure_oracle_client_initialized() -> None:
        """oracledb thick モード (Oracle Client 経由) を idempotent に初期化する。"""

        try:
            import oracledb  # noqa: PLC0415
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "oracledb is required for mcframe (Oracle) connections: "
                "run pip install oracledb."
            ) from exc
        try:
            oracledb.init_oracle_client()
        except oracledb.ProgrammingError:
            # すでに初期化済みのケースは無視
            pass

    def _connect(self):
        self._ensure_oracle_client_initialized()
        import oracledb  # noqa: PLC0415

        return oracledb.connect(
            user=self._user, password=self._password, dsn=self._dsn
        )

    def _fetch_rows(self, sql_template: str) -> Iterator[tuple]:
        sql = sql_template.format_map(
            {"schema": self._schema, "co_cd": self._co_cd}
        )
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql)
                for row in cur:
                    yield row

    # -- BridgeErpService implementation -------------------------------------

    def ping(self) -> None:
        """Oracle へ接続して ``SELECT 1 FROM DUAL`` を投げる軽量チェック。"""

        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1 FROM DUAL")
                cur.fetchone()

    def fetch_integrated_records(
        self, *, upload: Any = None
    ) -> list[dict[str, Any]]:
        raw_rows = list(self._fetch_rows(_MASTER_SQL))
        rec_dicts = [dict(zip(INTEGRATED_HEADERS, row)) for row in raw_rows]
        return append_supplier_use_lines_after_inputs(rec_dicts)

    def fetch_item_rows(self, *, upload: Any = None) -> list[tuple]:
        return list(self._fetch_rows(_ITEM_TABLE_SQL))

    def fetch_order_rows(self, *, upload: Any = None) -> list[tuple]:
        return list(self._fetch_rows(_ORDER_TABLE_SQL))

    def fetch_resource_rows(
        self,
        *,
        upload: Any = None,
        sort_order_map: dict[str, int] | None = None,
    ) -> list[tuple]:
        base_rows = list(self._fetch_rows(_RESOURCE_TABLE_SQL))
        sort_map = sort_order_map or {}
        out: list[tuple] = []
        for line_cd, line_nm, resource_grp in base_rows:
            sort_order = sort_map.get(str(line_cd or "").strip())
            out.append(
                (
                    line_cd,
                    line_nm,
                    resource_grp,
                    "" if sort_order is None else sort_order,
                )
            )
        return out

    def fetch_inventory_rows(self, *, upload: Any = None) -> list[tuple]:
        return list(self._fetch_rows(_INVENTORY_TABLE_SQL))
