"""顧客別ロジックの基底クラスおよびデータ構造。

新しい顧客固有の挙動を追加する場合は ``CustomerStrategy`` を継承した
クラスを ``core.customers.<id>`` モジュールに作成し、
``core.customers.__init__`` のレジストリに登録するだけで済む構成にしてある。

設計意図:
    - ``apps/viewer`` や ``apps/bridge`` 配下のハンドラに散らばっていた
      ``if customer_id == "phc"`` のようなハードコード分岐を廃止し、
      Strategy パターンで顧客モジュール側へ寄せる。
    - 既定挙動 (``DefaultCustomer``) を提供することで、
      未登録顧客や顧客未選択時にも安全に動く。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any, Callable, Iterable, Optional


# ---------------------------------------------------------------------------
# Generic value objects shared across apps and templates
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PsiRowDefinition:
    """PSI ビューの 1 行 (Supply / Demand / Stock など) の定義。"""

    row_type: str  # 'Supply' | 'Demand' | 'Stock'
    type_main: str  # 行頭ラベル (例: 'Supply', 'P', 'I')
    type_sub: str  # サブラベル (PHC の 'Export'/'Internal' など)
    type_rowspan: int  # 行頭セルの rowspan (0 のときは main 列を出さない)
    customer_bucket: Optional[str]  # split 集計時のバケット (None なら通常集計)


@dataclass
class BridgeButton:
    """Bridge トップ画面の TRANSACTION ボタン定義。"""

    kind: str  # 'order' | 'prd_plan' | 'inventory' | 'inventory_wip'
    label: str
    disabled: bool = False


@dataclass
class CustomerView:
    """テンプレートに渡す顧客プレゼンテーション情報。

    テンプレート側から ``customer_view.flags.psi_split_by_customer`` のように
    アクセスできるよう、フラグはネストした dict 構造を持つ。
    """

    id: str
    label: str
    flags: dict[str, Any] = field(default_factory=dict)
    css: dict[str, str] = field(default_factory=dict)
    bridge_buttons: list[BridgeButton] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Strategy interface
# ---------------------------------------------------------------------------


class CustomerStrategy:
    """顧客固有の挙動を集約する抽象クラス。

    既定実装はすべて『何もしない / フォールバックさせる』方向に倒してあるので、
    新顧客モジュールでは差分を持つメソッドだけオーバーライドすればよい。
    """

    #: 顧客 ID (lower-case)。レジストリのキーと一致させる。
    id: str = ""
    #: UI 表示用ラベル。
    label: str = ""

    # ------------------------------------------------------------------
    # PSI view (Viewer)
    # ------------------------------------------------------------------

    def psi_split_by_customer(self) -> bool:
        """PSI を Customer (Export/Internal) で行分割するか。"""

        return False

    def psi_table_extra_class(self) -> Optional[str]:
        """PSI テーブルに付与する追加 CSS クラス名。"""

        return None

    def psi_row_definitions(self) -> list[PsiRowDefinition]:
        """PSI ビューの行構成。"""

        return [
            PsiRowDefinition("Supply", "Supply", "", 1, None),
            PsiRowDefinition("Demand", "Demand", "", 1, None),
            PsiRowDefinition("Stock", "Stock", "", 1, None),
        ]

    def build_psi_split_aggs(
        self,
        start_date: date,
        end_date: date,
        db_factory: Callable[[], Any],
    ) -> tuple[dict, dict]:
        """Customer 軸で分けた supply / demand 集計を返す。

        Returns:
            (supply_split, demand_split) のタプル。
            - supply_split[(item, day, bucket, machine)] -> qty
            - demand_split[(item, day, bucket, machine)] -> qty
        既定では split を行わないため、空 dict を返す。
        """

        return ({}, {})

    # ------------------------------------------------------------------
    # Bridge Excel loader
    # ------------------------------------------------------------------

    def load_excel_export_rows(
        self, kind: str, raw: bytes
    ) -> Optional[list[tuple]]:
        """顧客固有の Excel ローダを呼ぶ。None を返した場合は共通処理を継続。"""

        return None

    # ------------------------------------------------------------------
    # Bridge optional features
    # ------------------------------------------------------------------

    def supports_monthly_result(self) -> bool:
        """``/monthly-result`` エンドポイントを表示・受付するか。"""

        return False

    def parse_monthly_result(self, raw: bytes) -> list[dict[str, str]]:
        """Monthly Result Excel のバイト列を表示用 dict 行に解析する。

        ``supports_monthly_result()`` が False の顧客では呼ばれない想定。
        """

        raise NotImplementedError(
            "Monthly Result is not supported for this customer."
        )

    # ------------------------------------------------------------------
    # Bridge UI
    # ------------------------------------------------------------------

    def bridge_buttons(self) -> list[BridgeButton]:
        """Bridge トップの TRANSACTION ボタン構成。"""

        return [
            BridgeButton("order", "Order"),
            BridgeButton("inventory", "Inventory"),
        ]

    # ------------------------------------------------------------------
    # Viewer header menu (Schedule Viewer)
    # ------------------------------------------------------------------

    def viewer_show_daily_schedule(self) -> bool:
        """Viewer ドロップダウンに Daily Schedule を出すか。"""

        return True

    def viewer_show_psi_viewer(self) -> bool:
        """Viewer ドロップダウンに PSI Viewer を出すか。"""

        return True

    # ------------------------------------------------------------------
    # Template adapter
    # ------------------------------------------------------------------

    def to_view(self) -> CustomerView:
        """テンプレートで参照しやすい構造体に整形して返す。"""

        return CustomerView(
            id=self.id,
            label=self.label,
            flags={
                "psi_split_by_customer": self.psi_split_by_customer(),
                "viewer_show_daily_schedule": self.viewer_show_daily_schedule(),
                "viewer_show_psi_viewer": self.viewer_show_psi_viewer(),
            },
            css={
                "psi_table_extra_class": self.psi_table_extra_class() or "",
            },
            bridge_buttons=list(self.bridge_buttons()),
        )


class DefaultCustomer(CustomerStrategy):
    """未登録顧客 / 顧客未選択時のフォールバック。"""

    id = ""
    label = ""


# ---------------------------------------------------------------------------
# Helpers shared between concrete strategies
# ---------------------------------------------------------------------------


def iter_buttons(buttons: Iterable[BridgeButton]) -> list[BridgeButton]:
    """テンプレート側で iterate しやすいよう list に正規化するユーティリティ。"""

    return list(buttons)
