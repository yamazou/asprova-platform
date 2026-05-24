"""PEB Item Table builder tests."""

from __future__ import annotations

import io

from openpyxl import Workbook

from core.erp.excel.peb_exports import (
    PEB_ITEM_TABLE_HEADERS,
    load_peb_item_table_rows_from_xlsx_bytes,
)


def _mini_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "master code"
    ws.append(
        [
            "FGcode",
            "Model",
            "process",
            "Infusion code",
            "Ink Bag",
            "Ink Code",
            "Ink Description",
        ]
    )
    ws.append(
        [
            "TESTFG001",
            41418223,
            "Main line",
            "INF001",
            2,
            611027300,
            "INK,991CGP",
        ]
    )
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_item_table_maps_master_code_specs() -> None:
    rows = load_peb_item_table_rows_from_xlsx_bytes(_mini_workbook_bytes())
    assert len(rows) == 1
    assert rows[0] == (
        "TESTFG001",
        "",
        "41418223",
        "INK,991CGP",
        "611027300",
        "2",
    )
    assert PEB_ITEM_TABLE_HEADERS[2] == "Item_Spec1Code"
