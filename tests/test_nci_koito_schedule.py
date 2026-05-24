"""NCI KOITO delivery schedule builder tests."""

from __future__ import annotations

import csv
import io
from pathlib import Path

from openpyxl import Workbook
from datetime import datetime

from core.erp.excel.nci_koito_schedule import (
    KOITO_CUSTOMER,
    build_koito_delivery_schedule_rows,
)

_REFERENCE_XLSX = Path(
    r"c:\Users\lenovo\OneDrive\0_BAHTERA_WORK\NCI\Prototype\data\Delivery date_②_May.xlsx"
)
_REFERENCE_CSV = Path(
    r"c:\Users\lenovo\OneDrive\0_BAHTERA_WORK\NCI\Prototype\1\delivery_schedule_KOITO.csv"
)


def _mini_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet 1"
    ws.append([None, None, None, None, None, None])
    ws.append(
        [None, "KODE KANBAN", "PART", "NAME", "Item CD", "SNP", datetime(2026, 5, 4)]
    )
    ws.append(
        [None, "S889", "10046-8M004", "BODY", "NCI_TEST_ITEM", 4, 28]
    )
    # header row must include KODE KANBAN (row above uses Item CD in col E)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_build_koito_rows_from_mini_workbook() -> None:
    rows = build_koito_delivery_schedule_rows(_mini_workbook_bytes())
    assert rows == [
        ("KOITO00001", "NCI_TEST_ITEM", "5/4/2026", "28", "S889", KOITO_CUSTOMER),
    ]


def test_matches_reference_csv_when_files_present() -> None:
    if not _REFERENCE_XLSX.is_file() or not _REFERENCE_CSV.is_file():
        return
    built = build_koito_delivery_schedule_rows(_REFERENCE_XLSX.read_bytes())
    with _REFERENCE_CSV.open(encoding="utf-8-sig", newline="") as f:
        ref = [tuple(r) for r in csv.reader(f)][1:]
    assert len(built) == len(ref)
    assert built[0][0] == "KOITO00001"
    assert built[-1][0] == f"KOITO{len(built):05d}"
    assert [r[0] for r in built] == [
        f"KOITO{i:05d}" for i in range(1, len(built) + 1)
    ]
    assert {tuple(r[1:6]) for r in built} == {tuple(r[1:6]) for r in ref}
