"""SAP B1 (SQL Server) 用 Bridge サービス実装。

Bridge アプリは ``apps/bridge/app.py`` から ``SapB1BridgeService`` を生成して
利用する。``connect_sqlserver`` を直接呼ぶのは本サービス内のみ。
"""

from __future__ import annotations

from typing import Any

from core.erp._base import BridgeErpService, NotSupportedError

from .connection import connect_sqlserver
from .integrated_master import fetch_integrated_master_rows_from_sqlserver
from .inventory_table import fetch_inventory_rows_from_sqlserver
from .item_table import fetch_item_table_rows_from_sqlserver
from .order_table import fetch_order_rows_from_sqlserver


class SapB1BridgeService(BridgeErpService):
    """SAP Business One (SQL Server) 用 Bridge サービス。

    Args:
        server / database / user / password: SQL Server 接続情報
        timeout: 接続タイムアウト秒
    """

    def __init__(
        self,
        *,
        server: str,
        database: str,
        user: str,
        password: str,
        timeout: int = 30,
    ) -> None:
        if not (user and password and server and database):
            raise RuntimeError(
                "Not connected. Please connect to the database from Connect first."
            )
        self._server = server
        self._database = database
        self._user = user
        self._password = password
        self._timeout = timeout

    def _connect(self, *, timeout: int | None = None):
        return connect_sqlserver(
            self._server,
            self._database,
            self._user,
            self._password,
            timeout=timeout if timeout is not None else self._timeout,
        )

    # -- BridgeErpService implementation -------------------------------------

    def ping(self) -> None:
        """SQL Server へ短いタイムアウトで接続して ``SELECT 1`` を投げる。"""

        conn = self._connect(timeout=3)
        try:
            cur = conn.cursor()
            try:
                cur.execute("SELECT 1")
                cur.fetchone()
            finally:
                cur.close()
        finally:
            conn.close()

    def fetch_integrated_records(
        self, *, upload: Any = None
    ) -> list[dict[str, Any]]:
        conn = self._connect()
        try:
            return fetch_integrated_master_rows_from_sqlserver(conn)
        finally:
            conn.close()

    def fetch_item_rows(self, *, upload: Any = None) -> list[tuple]:
        conn = self._connect()
        try:
            return fetch_item_table_rows_from_sqlserver(conn)
        finally:
            conn.close()

    def fetch_order_rows(self, *, upload: Any = None) -> list[tuple]:
        conn = self._connect()
        try:
            return fetch_order_rows_from_sqlserver(conn)
        finally:
            conn.close()

    def fetch_inventory_rows(self, *, upload: Any = None) -> list[tuple]:
        conn = self._connect()
        try:
            return fetch_inventory_rows_from_sqlserver(conn)
        finally:
            conn.close()

    def fetch_resource_rows(
        self,
        *,
        upload: Any = None,
        sort_order_map: dict[str, int] | None = None,
    ) -> list[tuple]:
        raise NotSupportedError(
            "This output is not supported for SAP Business One connections "
            "because it uses mcframe / Oracle tables."
        )
