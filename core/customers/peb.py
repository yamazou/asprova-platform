"""PEB 顧客向けの固有ロジック。

主な差分:
    - Excel ローダ (Order/Prd Plan/Inventory/Inventory WIP) を専用関数に切替
    - Bridge トップのボタン構成 (Shipping / Prd Plan / Inv. FG / Inv. WIP)
"""

from __future__ import annotations

from typing import Optional

from core.erp.excel.peb_exports import (
    load_peb_inventory_rows_from_xlsx_bytes,
    load_peb_inventory_wip_rows_from_xlsx_bytes,
    load_peb_monthly_result_rows_from_xlsx_bytes,
    load_peb_order_rows_from_xlsx_bytes,
    load_peb_prd_plan_rows_from_xlsx_bytes,
)

from .base import BridgeButton, CustomerStrategy


class PebCustomer(CustomerStrategy):
    id = "peb"
    label = "PEB"

    # ------------------------------------------------------------------
    # Viewer header menu (PEB 納品: Gantt + Monthly Result のみ)
    # ------------------------------------------------------------------

    def viewer_show_daily_schedule(self) -> bool:
        return False

    def viewer_show_psi_viewer(self) -> bool:
        return False

    # ------------------------------------------------------------------
    # Bridge Excel loader
    # ------------------------------------------------------------------

    _EXCEL_LOADERS = {
        "order": load_peb_order_rows_from_xlsx_bytes,
        "prd_plan": load_peb_prd_plan_rows_from_xlsx_bytes,
        "inventory": load_peb_inventory_rows_from_xlsx_bytes,
        "inventory_wip": load_peb_inventory_wip_rows_from_xlsx_bytes,
    }

    def load_excel_export_rows(
        self, kind: str, raw: bytes
    ) -> Optional[list[tuple]]:
        loader = self._EXCEL_LOADERS.get(kind)
        if loader is None:
            return None
        return loader(raw)

    # ------------------------------------------------------------------
    # Bridge UI
    # ------------------------------------------------------------------

    def bridge_buttons(self) -> list[BridgeButton]:
        return [
            BridgeButton("order", "Shipping"),
            BridgeButton("prd_plan", "Prd Plan"),
            BridgeButton("inventory", "Inv. FG"),
            BridgeButton("inventory_wip", "Inv. WIP"),
        ]

    # ------------------------------------------------------------------
    # Monthly Result (PEB 専用)
    # ------------------------------------------------------------------

    def supports_monthly_result(self) -> bool:
        return True

    def parse_monthly_result(self, raw: bytes) -> list[dict[str, str]]:
        return load_peb_monthly_result_rows_from_xlsx_bytes(raw)
