"""Bridge アプリ向けの ERP 別データ取得サービス (抽象基底)。

設計意図:
    - ``apps/bridge/app.py`` は HTTP / セッション / CSV 化など
      Web 層の関心事だけを持ち、ERP 固有の SQL や Excel 解析は
      この interface 配下の各サブパッケージへ閉じ込める。
    - 顧客納品時に不要な ERP サブパッケージ (例: PHC 納品では
      ``core/erp/sap_b1`` 一式) を物理削除しても import で壊れないよう、
      呼び出し側は lazy import すること。
    - 各 ERP に存在しない機能 (例: SAP B1 では Resource Table、
      mcframe では Inv. WIP) は ``NotSupportedError`` を送出する。
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Optional


class NotSupportedError(RuntimeError):
    """この ERP では未対応の操作。"""


class BridgeErpService(ABC):
    """Bridge アプリで使う ERP 抽象アダプタ。

    実装は各 ``core/erp/<system>/service.py`` に置く。コンストラクタで
    必要な接続情報や顧客プロファイルを受け取り、メソッド単位に必要なときだけ
    実接続を張る (毎回 close する) のが原則。
    """

    # ------------------------------------------------------------------
    # Optional: 接続生存確認 (UI の CONNECTED 表示用)
    # ------------------------------------------------------------------

    def ping(self) -> None:
        """ERP との接続を軽量に検証する。

        既定実装は何もしない (Excel など接続を伴わない ERP 用)。
        DB 系の実装では短いタイムアウトで実接続を試み、失敗時は例外を送出すること。
        """
        return None

    # ------------------------------------------------------------------
    # Required: 全 ERP で取得対象となる行
    # ------------------------------------------------------------------

    @abstractmethod
    def fetch_integrated_records(
        self, *, upload: Any = None
    ) -> list[dict[str, Any]]:
        """Integrated Master を ``INTEGRATED_HEADERS`` に対応する dict 行で返す。

        Cycle Time の差し込みなど ERP 横断の後段処理を呼び出し側でかけられるよう、
        tuple ではなく dict を返す。
        """

    @abstractmethod
    def fetch_item_rows(self, *, upload: Any = None) -> list[tuple]:
        """Item Table の tuple 行 (列順は ``ITEM_TABLE_HEADERS`` と一致)。"""

    @abstractmethod
    def fetch_order_rows(self, *, upload: Any = None) -> list[tuple]:
        """Order Table の tuple 行 (列順は ``ORDER_TABLE_HEADERS`` と一致)。"""

    # ------------------------------------------------------------------
    # Optional: ERP によっては未対応のもの
    # ------------------------------------------------------------------

    def fetch_resource_rows(
        self,
        *,
        upload: Any = None,
        sort_order_map: Optional[dict[str, int]] = None,
    ) -> list[tuple]:
        raise NotSupportedError(
            "Resource Table is not supported for this ERP."
        )

    def fetch_inventory_rows(self, *, upload: Any = None) -> list[tuple]:
        raise NotSupportedError(
            "Inventory Table is not supported for this ERP."
        )

    def fetch_prd_plan_rows(self, *, upload: Any = None) -> list[tuple]:
        raise NotSupportedError(
            "Prd Plan is supported only for Excel import."
        )

    def fetch_inventory_wip_rows(self, *, upload: Any = None) -> list[tuple]:
        raise NotSupportedError(
            "Inv. WIP is supported only for Excel import."
        )
