"""NCI HPM delivery schedule builder tests."""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from core.erp.excel.nci_hpm_schedule import (
    build_hpm_delivery_schedule_rows,
)

_REFERENCE_XLSX = Path(
    r"c:\Users\lenovo\OneDrive\0_BAHTERA_WORK\NCI\Prototype\data\Delivery date_①_May.xlsx"
)

def _mini_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "BRIO PLANT 2"
    ws.append([None] * 6)
    ws.append([None, "BRIO Delivery Plant 2"])
    ws.append([None] * 6)
    ws.append(
        [
            None,
            "PART CODE",
            "PART NAME",
            "mcframe Item code",
            "COLOR",
            "MODEL",
            datetime(2026, 5, 4),
        ]
    )
    ws.append(
        [
            None,
            "71103TG4 K500",
            "GRILLE",
            "NCI_TG4_71103_K500_0_NH696L_A",
            None,
            "3UB RH",
            180,
        ]
    )
    ws.append([None, None, None, None, None, "3UB RH IN"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_build_hpm_rows_model_with_own_qty_only() -> None:
    rows = build_hpm_delivery_schedule_rows(_mini_workbook_bytes())
    assert len(rows) == 1
    assert rows[0] == (
        "HPM00001",
        "NCI_TG4_71103_K500_0_NH696L_A",
        "5/4/2026",
        "180",
        "3UB RH",
        "BRIO PLANT 2",
    )


def test_reference_workbook_schedule_rows_only() -> None:
    if not _REFERENCE_XLSX.is_file():
        return
    built = build_hpm_delivery_schedule_rows(_REFERENCE_XLSX.read_bytes())
    assert len(built) == 1258
    assert built[0][0] == "HPM00001"
    for i, row in enumerate(built, start=1):
        assert row[0] == f"HPM{i:05d}"
    customers = {r[5] for r in built}
    assert customers == {
        "BRIO PLANT 2",
        "BR-V",
        "HR-V & HR-V EXP",
        "WR-V + WR-V EXP",
    }
    assert " IN" not in "".join(r[4] for r in built)


def test_color_sheet_uses_color_not_dest() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "BR-V"
    ws.append(["PART CODE", "PART NAME", "mcframe Item code", "COLOR", "DEST", datetime(2026, 5, 1)])
    ws.append(
        [
            "X",
            "PART A",
            "NCI_TEST_ITEM_1",
            "BLACK",
            "BRV RH",
            10,
        ]
    )
    ws.append([None, None, None, "BLACK", "BRV RH EXP", 5])
    ws.append([None, None, None, None, "TOTAL", 15])
    buf = io.BytesIO()
    wb.save(buf)
    rows = build_hpm_delivery_schedule_rows(buf.getvalue())
    assert len(rows) == 2
    assert rows[0][4] == "BLACK"
    assert rows[1][4] == "BLACK"
    assert rows[0][5] == "BR-V"
