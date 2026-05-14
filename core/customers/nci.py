"""NCI 顧客向けの固有ロジック。

Data Bridge の TRANSACTION で Order / Inventory を利用しないため、
ボタンは表示したままグレーアウト（無効）にする。
"""

from __future__ import annotations

from .base import BridgeButton, CustomerStrategy


class NciCustomer(CustomerStrategy):
    id = "nci"
    label = "NCI"

    def bridge_buttons(self) -> list[BridgeButton]:
        return [
            BridgeButton("order", "Order", disabled=True),
            BridgeButton("inventory", "Inventory", disabled=True),
        ]
