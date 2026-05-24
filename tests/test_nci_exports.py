"""NCI Excel master export tests."""

from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook

from core.erp.excel.nci_exports import (
    build_nci_integrated_master_records,
    load_nci_item_table_rows_from_xlsx_bytes,
    load_nci_resource_table_rows_from_xlsx_bytes,
)

_DATA = Path(r"c:\Users\lenovo\OneDrive\0_BAHTERA_WORK\NCI\Prototype\data")
_BOM = _DATA / "BOM.xlsx"
_ITEM_LINE = _DATA / "ItemLine.xlsx"
_ITEM = _DATA / "Item.xlsx"
_LINE = _DATA / "Line.xlsx"


def test_reference_integrated_bom_and_itemline() -> None:
    if not _BOM.is_file() or not _ITEM_LINE.is_file():
        return
    records = build_nci_integrated_master_records(
        _BOM.read_bytes(),
        _ITEM_LINE.read_bytes(),
    )
    assert len(records) > 1000
    inst_types = {r["INST_TYP"] for r in records}
    assert "I" in inst_types
    assert "U" in inst_types
    prod = [r for r in records if r.get("ITM_RESOURCE") == "PROD"]
    assert not prod


def test_reference_item_and_resource_tables() -> None:
    if not _ITEM.is_file() or not _LINE.is_file():
        return
    items = load_nci_item_table_rows_from_xlsx_bytes(_ITEM.read_bytes())
    assert len(items) > 1000
    assert items[0][0]
    resources = load_nci_resource_table_rows_from_xlsx_bytes(_LINE.read_bytes())
    assert len(resources) >= 5
    assert resources[0][0] == "A001"
    m_line = next(r for r in resources if r[0] == "M001")
    assert m_line[2] == "INJECTION"


def test_mini_bom_itemline_integrated_pair() -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "Parent item CD",
            "Item name",
            "Model No.",
            "BOM pattern",
            "Use start date",
            "Use end date",
            "Operation CD",
            "Operation name",
            "Child item CD",
            "Child input qty.",
        ]
    )
    ws.append(
        [
            "PARENT1",
            "Parent",
            None,
            1,
            None,
            None,
            "001",
            "Op",
            "CHILD1",
            2,
        ]
    )
    bom_buf = io.BytesIO()
    wb.save(bom_buf)

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.append(
        [
            "Item CD",
            "Item name",
            "Line CD",
            "Line name",
            "Main line",
            "Standard load",
            "Cycle time",
        ]
    )
    ws2.append(["PARENT1", "Parent", "A002", "Assy Other", "true", 30, None])
    il_buf = io.BytesIO()
    wb2.save(il_buf)

    records = build_nci_integrated_master_records(bom_buf.getvalue(), il_buf.getvalue())
    assert any(
        r["P_ITM_CD"] == "PARENT1"
        and r["INST_TYP"] == "I"
        and r["ITM_RESOURCE"] == "CHILD1"
        for r in records
    )
    u_row = next(
        r
        for r in records
        if r["P_ITM_CD"] == "PARENT1" and r["INST_TYP"] == "U" and r["ITM_RESOURCE"] == "A002"
    )
    assert u_row["PRODUCTION"] == "30mp"
