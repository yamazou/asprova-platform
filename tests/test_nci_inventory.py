"""NCI Inventory.xlsx import tests."""

from __future__ import annotations

import io
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from core.erp.excel.nci_exports import (
    _adjust_nci_allocatable_qty,
    load_nci_inventory_rows_from_xlsx_bytes,
)
from core.erp.inventory_aggregate import aggregate_inventory_rows_by_itm_cd

_PROTOTYPE_INVENTORY = Path(
    r"c:\Users\lenovo\OneDrive\0_BAHTERA_WORK\NCI\Prototype\data\Inventory.xlsx"
)


def _inventory_bytes(rows: list[list], *, sheet: str = "Current stock list") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_adjust_qty_subtracts_while_at_least_100000() -> None:
    assert _adjust_nci_allocatable_qty(Decimal("100123.757")) == Decimal("123.757")
    assert _adjust_nci_allocatable_qty(Decimal("101082.667")) == Decimal("1082.667")
    assert _adjust_nci_allocatable_qty(Decimal("100000")) == Decimal(0)
    assert _adjust_nci_allocatable_qty(Decimal("200000")) == Decimal(0)


def test_inventory_sums_duplicate_item_rows() -> None:
    raw = _inventory_bytes(
        [
            ["Item CD", "Item name", "Allocatable qty."],
            ["1010101007", "A", 100000],
            ["1010101007", "A", 100000],
            ["1010101002", "B", 100123.757],
        ]
    )
    rows = aggregate_inventory_rows_by_itm_cd(
        load_nci_inventory_rows_from_xlsx_bytes(raw)
    )
    by_item = {r[1]: r[2] for r in rows}
    assert by_item["1010101007"] == "0"
    assert by_item["1010101002"] == "123.757"
    assert rows[0][0] == "INV00001"


def test_prototype_inventory_sample_items() -> None:
    if not _PROTOTYPE_INVENTORY.is_file():
        return
    raw = _PROTOTYPE_INVENTORY.read_bytes()
    rows = aggregate_inventory_rows_by_itm_cd(
        load_nci_inventory_rows_from_xlsx_bytes(raw)
    )
    by_item = {r[1]: r[2] for r in rows}
    assert by_item["1010101002"] == "123.757"
    assert by_item["1010101005"] == "1082.667"
    assert by_item["1010101006"] == "0"
    assert by_item["1010101007"] == "0"
    assert len(rows) >= 1800
