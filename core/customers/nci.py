"""NCI 顧客向けの固有ロジック。

Data Bridge は Excel ソース（Oracle 接続なし）。
MASTER: BOM + ItemLine + Cycle Time Master → Integrated Master、
Item master → Item Table、Line master → Resource Table。
TRANSACTION: KOITO / HPM / Inventory（Inventory.xlsx）。
"""

from __future__ import annotations

from typing import Optional

from core.erp.excel.nci_exports import (
    NCI_ITEM_TABLE_HEADERS,
    build_nci_integrated_master_records,
    load_nci_inventory_rows_from_xlsx_bytes,
    load_nci_item_table_rows_from_xlsx_bytes,
    load_nci_resource_table_rows_from_xlsx_bytes,
)

from .base import BridgeButton, CustomerStrategy

_INVENTORY_MARGIN_START_PX = 20


class NciCustomer(CustomerStrategy):
    id = "nci"
    label = "NCI"

    _EXCEL_LOADERS = {
        "item": load_nci_item_table_rows_from_xlsx_bytes,
        "resource": load_nci_resource_table_rows_from_xlsx_bytes,
        "inventory": load_nci_inventory_rows_from_xlsx_bytes,
    }

    def load_excel_export_rows(
        self, kind: str, raw: bytes
    ) -> Optional[list[tuple]]:
        loader = self._EXCEL_LOADERS.get(kind)
        if loader is None:
            return None
        if kind == "resource":
            return loader(raw, sort_order_map=None)
        return loader(raw)

    def item_table_csv_headers(self) -> tuple[str, ...]:
        return NCI_ITEM_TABLE_HEADERS

    def bridge_master_buttons(self) -> list[BridgeButton]:
        return [
            BridgeButton("integrated", "Integrated Master"),
            BridgeButton("rl_output", "R/L Output"),
            BridgeButton("item", "Item Table"),
            BridgeButton("resource", "Resource Table"),
            BridgeButton("line_cycle", "Cycle Time Master"),
        ]

    def bridge_buttons(self) -> list[BridgeButton]:
        return [
            BridgeButton("koito", "KOITO"),
            BridgeButton("hpm", "HPM"),
            BridgeButton(
                "inventory",
                "Inventory",
                margin_start_px=_INVENTORY_MARGIN_START_PX,
            ),
        ]


def build_nci_integrated_from_uploads(
    bom_raw: bytes,
    item_line_raw: bytes,
) -> list[dict[str, str]]:
    """Bridge から BOM / ItemLine のバイト列を渡して Integrated Master を生成する。"""
    return build_nci_integrated_master_records(bom_raw, item_line_raw)
