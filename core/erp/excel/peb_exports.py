from __future__ import annotations

import io
import re
from datetime import datetime


def _normalize_header_name(raw: object) -> str:
    text = "" if raw is None else str(raw).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def _format_excel_date(value: object) -> str:
    if isinstance(value, datetime):
        return f"{value.month}/{value.day}/{value.year}"
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(text, fmt)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except ValueError:
            continue
    return text


def load_peb_order_rows_from_xlsx_bytes(raw: bytes) -> list[tuple[str, str, str, str, str]]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "openpyxl is required to read Excel (.xlsx) files."
        ) from exc
    wb = load_workbook(filename=io.BytesIO(raw), read_only=True, data_only=True)
    try:
        if "Shipment Plan" not in wb.sheetnames:
            raise RuntimeError("PEB order import requires a 'Shipment Plan' sheet.")
        ws = wb["Shipment Plan"]
        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if not header_row:
            return []

        header_index = {_normalize_header_name(v): i for i, v in enumerate(header_row)}

        def _idx(*names: str) -> int | None:
            for n in names:
                idx = header_index.get(_normalize_header_name(n))
                if idx is not None:
                    return idx
            return None

        idx_item = _idx("Item Code", "ITM_CD")
        idx_qty = _idx("Qty", "REQ_QTY")
        idx_dlv = _idx("Exfact Date", "DLV_DT")
        idx_req = _idx("Index", "REQ_NO")
        idx_cust = _idx("Customer Name", "CUST_CD")

        if idx_item is None or idx_qty is None or idx_dlv is None:
            raise RuntimeError(
                "Shipment Plan is missing required columns. "
                "Required: Item Code, Qty, Exfact Date."
            )

        out: list[tuple[str, str, str, str, str]] = []
        seq = 1
        for values in rows:
            item = str(values[idx_item] if idx_item < len(values) else "").strip()
            qty_raw = values[idx_qty] if idx_qty < len(values) else ""
            dlv_raw = values[idx_dlv] if idx_dlv < len(values) else ""
            if not item:
                continue

            qty_text = "" if qty_raw is None else str(qty_raw).strip()
            if not qty_text:
                continue

            req_raw = values[idx_req] if (idx_req is not None and idx_req < len(values)) else ""
            req_no = str(req_raw).strip() if req_raw is not None else ""
            if not req_no:
                req_no = f"PEB-{seq}"

            cust_raw = values[idx_cust] if (idx_cust is not None and idx_cust < len(values)) else ""
            cust_cd = str(cust_raw).strip() if cust_raw is not None else ""

            out.append((req_no, item, _format_excel_date(dlv_raw), qty_text, cust_cd))
            seq += 1
        return out
    finally:
        wb.close()


def load_peb_inventory_rows_from_xlsx_bytes(raw: bytes) -> list[tuple[str, str, str, str]]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "openpyxl is required to read Excel (.xlsx) files."
        ) from exc
    wb = load_workbook(filename=io.BytesIO(raw), read_only=True, data_only=True)
    try:
        if "FG stock" not in wb.sheetnames:
            raise RuntimeError("PEB inventory import requires an 'FG stock' sheet.")
        ws = wb["FG stock"]
        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if not header_row:
            return []

        item_idx = None
        tran_idx = None
        qty_idx = None
        for i, v in enumerate(header_row):
            key = _normalize_header_name(v)
            if key in ("itemdesc", "itemcode", "itmcd", "itm_cd"):
                item_idx = i
            elif key in ("tran", "transaction"):
                tran_idx = i
            elif key in ("qty", "quantity", "stkqty", "stockqty"):
                qty_idx = i
        if item_idx is None:
            item_idx = 0
        if qty_idx is None:
            qty_idx = 2 if len(header_row) > 2 else None

        if qty_idx is None:
            raise RuntimeError("No stock quantity column was found in FG stock.")

        inv_dt = _format_excel_date(header_row[qty_idx] if qty_idx < len(header_row) else "")

        out: list[tuple[str, str, str, str]] = []
        seq = 1
        for values in rows:
            item = str(values[item_idx] if item_idx < len(values) else "").strip()
            if not item:
                continue
            tran_val = str(values[tran_idx] if (tran_idx is not None and tran_idx < len(values)) else "").strip().lower()
            if tran_val and tran_val not in ("on hand", "onhand", "stock"):
                continue
            qty_raw = values[qty_idx] if qty_idx < len(values) else ""
            qty_text = "" if qty_raw is None else str(qty_raw).strip()
            if not qty_text:
                continue
            inv_cd = f"INV{seq:05d}"
            out.append((inv_cd, item, qty_text, inv_dt))
            seq += 1
        return out
    finally:
        wb.close()


def load_peb_inventory_wip_rows_from_xlsx_bytes(raw: bytes) -> list[tuple[str, str, str, str]]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "openpyxl is required to read Excel (.xlsx) files."
        ) from exc
    wb = load_workbook(filename=io.BytesIO(raw), read_only=True, data_only=True)
    try:
        ws_name = next((n for n in wb.sheetnames if str(n).startswith("Stock_list_")), None)
        if not ws_name:
            raise RuntimeError("PEB WIP inventory import requires a sheet starting with 'Stock_list_'.")
        ws = wb[ws_name]
        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if not header_row:
            return []

        item_idx = None
        qty_idx = None
        for i, v in enumerate(header_row):
            key = _normalize_header_name(v)
            if key in ("material", "itemdesc", "itemcode", "itmcd", "itm_cd"):
                item_idx = i
            elif key in ("available", "qty", "quantity", "stkqty", "stockqty"):
                qty_idx = i
        if item_idx is None:
            item_idx = 0
        if qty_idx is None:
            qty_idx = 2 if len(header_row) > 2 else None
        if qty_idx is None:
            raise RuntimeError("No stock quantity column was found in the Stock_list_ sheet.")

        inv_dt = ""
        m = re.match(r"^Stock_list_(\d{4})(\d{2})(\d{2})", ws_name)
        if m:
            inv_dt = f"{int(m.group(2))}/{int(m.group(3))}/{int(m.group(1))}"

        out: list[tuple[str, str, str, str]] = []
        seq = 1
        for values in rows:
            item = str(values[item_idx] if item_idx < len(values) else "").strip()
            if not item:
                continue
            qty_raw = values[qty_idx] if qty_idx < len(values) else ""
            qty_text = "" if qty_raw is None else str(qty_raw).strip()
            if not qty_text:
                continue
            inv_cd = f"WIP{seq:05d}"
            out.append((inv_cd, item, qty_text, inv_dt))
            seq += 1
        return out
    finally:
        wb.close()


def load_peb_prd_plan_rows_from_xlsx_bytes(raw: bytes) -> list[tuple[str, str, str, str, str]]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "openpyxl is required to read Excel (.xlsx) files."
        ) from exc
    wb = load_workbook(filename=io.BytesIO(raw), read_only=True, data_only=True)
    try:
        ws_name = next((n for n in wb.sheetnames if str(n).startswith("ProductionPlan")), None)
        if not ws_name:
            raise RuntimeError("PEB Prd Plan import requires a 'ProductionPlan' sheet.")
        ws = wb[ws_name]
        rows = ws.iter_rows(values_only=True)
        _ = next(rows, None)  # title row
        header_row = next(rows, None)
        if not header_row:
            return []

        fg_idx = None
        tran_idx = None
        date_cols: list[tuple[int, str]] = []
        for i, v in enumerate(header_row):
            key = _normalize_header_name(v)
            if fg_idx is None and key in ("fgcode", "itemcode", "itmcd", "itm_cd"):
                fg_idx = i
            elif key in ("tran", "transaction"):
                tran_idx = i

            text = "" if v is None else str(v).strip()
            if not text:
                continue
            token = text.split()[0]
            dt = _format_excel_date(token)
            if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", dt):
                date_cols.append((i, dt))

        if fg_idx is None:
            raise RuntimeError("No FG Code column was found in the ProductionPlan sheet.")
        if not date_cols:
            raise RuntimeError("No date header column was found in the ProductionPlan sheet.")

        out: list[tuple[str, str, str, str, str]] = []
        seq = 1
        for values in rows:
            itm_cd = str(values[fg_idx] if fg_idx < len(values) else "").strip()
            if not itm_cd:
                continue
            tran_val = str(values[tran_idx] if (tran_idx is not None and tran_idx < len(values)) else "").strip().lower()
            if tran_val and tran_val != "production":
                continue

            for col_idx, dlv_dt in date_cols:
                qty_raw = values[col_idx] if col_idx < len(values) else ""
                qty_text = "" if qty_raw is None else str(qty_raw).strip()
                if not qty_text:
                    continue
                try:
                    if float(qty_text) == 0:
                        continue
                except ValueError:
                    pass
                req_no = f"PRDPLAN-{seq:06d}"
                out.append((req_no, itm_cd, dlv_dt, qty_text, ""))
                seq += 1
        return out
    finally:
        wb.close()


def _last_numeric_value(cells: tuple[object, ...]) -> float | None:
    for value in reversed(cells):
        if isinstance(value, (int, float)):
            return float(value)
    return None


def _format_numeric_text(value: float) -> str:
    if value.is_integer():
        return f"{int(value):,}"
    text = f"{value:,.3f}"
    return text.rstrip("0").rstrip(".")


def load_peb_monthly_result_rows_from_xlsx_bytes(raw: bytes) -> list[dict[str, str]]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "openpyxl is required to read Excel (.xlsx) files."
        ) from exc
    wb = load_workbook(filename=io.BytesIO(raw), read_only=True, data_only=True)
    try:
        if "Monthly_ym" not in wb.sheetnames:
            raise RuntimeError("PEB Monthly Result import requires a 'Monthly_ym' sheet.")
        ws = wb["Monthly_ym"]
        qty_by_line: dict[str, float] = {}
        wh_by_line: dict[str, float] = {}
        mode: str | None = None

        for row in ws.iter_rows(values_only=True):
            key_col2 = str(row[1] if len(row) > 1 and row[1] is not None else "").strip()
            key = str(row[2] if len(row) > 2 and row[2] is not None else "").strip()
            if not key:
                if key_col2 == "QTY":
                    mode = "qty"
                elif key_col2 == "WH":
                    mode = "wh"
                continue
            lowered = key.lower()
            if lowered == "sum of qty":
                mode = "qty-header"
                continue
            if lowered == "sum of working hour":
                mode = "wh-header"
                continue
            if key == "Grand Total":
                continue
            if mode not in ("qty", "wh"):
                continue

            total = _last_numeric_value(row[3:])
            if total is None:
                continue
            if mode == "qty":
                qty_by_line[key] = qty_by_line.get(key, 0.0) + total
            else:
                wh_by_line[key] = wh_by_line.get(key, 0.0) + total

        keys = sorted(set(qty_by_line.keys()) | set(wh_by_line.keys()))
        return [
            {
                "line_name": k,
                "production_qty": _format_numeric_text(qty_by_line.get(k, 0.0)),
                "working_hours": _format_numeric_text(wh_by_line.get(k, 0.0)),
            }
            for k in keys
        ]
    finally:
        wb.close()
