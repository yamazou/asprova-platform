"""
Recreate staging tables on SQL Server from SAP B1–style Excel (OITW_TMP, OITM_TMP, ORSG_TMP, BEG_INV, SHIP_SCH, …).

Reads the first worksheet, uses row 1 as column names, infers INT / DECIMAL / NVARCHAR
from the data, DROP + CREATE, then bulk INSERT.

Example (PowerShell):

  $env:SQLSERVER_PASSWORD = 'your-password'
  py -3 scripts/load_oitw_oitm_tmp_from_excel.py ^
    --server "HOST\\SQLEXPRESS" --database SW --user sa ^
    --oitw "C:\\Users\\you\\Downloads\\OITW_TMP.xlsx" ^
    --oitm "C:\\Users\\you\\Downloads\\OITM_TMP.xlsx"

  py -3 scripts/load_oitw_oitm_tmp_from_excel.py --no-oitw --no-oitm ^
    --orsg "C:\\Users\\you\\Downloads\\ORSG_TMP.xlsx" ...

  py -3 scripts/load_oitw_oitm_tmp_from_excel.py --no-oitw --no-oitm --no-orsg ^
    --beg-inv "C:\\Users\\you\\Downloads\\BEG_INV.xlsx" ...

  py -3 scripts/load_oitw_oitm_tmp_from_excel.py --no-oitw --no-oitm --no-orsg --no-beg-inv ^
    --ship-sch "C:\\Users\\you\\Downloads\\SHIP_SCH.xlsx" ...

  py -3 scripts/load_oitw_oitm_tmp_from_excel.py --dry-run --oitw ... --oitm ...
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from decimal import Decimal
from typing import Any, List, Sequence, Tuple

import numpy as np
import pandas as pd


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _coerce(value: Any, kind: str) -> Any:
    if _is_empty(value):
        return None
    if kind == "int":
        if isinstance(value, (int, float)):
            if isinstance(value, float) and (value != value or pd.isna(value)):
                return None
            return int(round(float(value)))
        s = str(value).strip()
        if s == "":
            return None
        return int(round(float(s)))
    if kind == "decimal":
        if isinstance(value, (int, float)):
            if isinstance(value, float) and (value != value or pd.isna(value)):
                return None
            return Decimal(str(value))
        s = str(value).strip()
        if s == "":
            return None
        return Decimal(s)
    # nvarchar
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.isoformat()
    if isinstance(value, dt.datetime):
        return value.isoformat()
    if isinstance(value, float):
        if value != value or pd.isna(value):
            return None
        if value == int(value):
            return str(int(value))
        return str(value).rstrip("0").rstrip(".") if "." in str(value) else str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, bool):
        return "Y" if value else "N"
    return str(value).strip() or None


def _bracket_ident(name: str) -> str:
    return "[" + str(name).replace("]", "]]") + "]"


def _name_suggests_decimal(col_name: str) -> bool:
    c = col_name.lower()
    keys = (
        "price",
        "qty",
        "quantity",
        "rate",
        "amount",
        "value",
        "weight",
        "volume",
        "factor",
        "commission",
        "cost",
        "balance",
        "percent",
        "in stock",
        "ordered",
        "minimum",
        "maximum",
        "multiple",
        "level",
        "limit",
        "days",
        "hours",
        "tolerance",
        "inventory value",
        "std cost",
        "assessable",
        "goods on hold",
        "pricing",
        "items per",
        "no. of items",
        "quantity per",
    )
    return any(k in c for k in keys) or "%" in c


def _name_suggests_code_text(col_name: str) -> bool:
    """SAP item / warehouse / account columns: keep alphanumeric codes even if Excel stores as number."""
    c = col_name.lower()
    keys = (
        " code",
        "account",
        "uom",
        "remarks",
        "description",
        "name",
        "text",
        "formula",
        "bar code",
        "serial",
        "picture",
        "property ",
        "currency",
        "vendor",
        "manufacturer",
        "group",
        "template",
        "definition",
        "method",
        "status",
        "warehouse",
        "bin",
        "line",
        "item no",
        "ncm ",
        "sac ",
        "cest ",
        "nve ",
        "tnved",
        "identification",
        "classification",
        "legal text",
        "source",
        " acct",
        "adj. acct",
    )
    if any(k in c for k in keys):
        return True
    if c.startswith("no.") or c.endswith(" no.") or " no." in c:
        return True
    return False


def _nvarchar_tier_from_series(s: pd.Series) -> str:
    lens = s.dropna().astype(str).str.len()
    mx = int(lens.max()) if len(lens) else 0
    if mx <= 128:
        return "NVARCHAR(256) NULL"
    if mx <= 512:
        return "NVARCHAR(512) NULL"
    if mx <= 4000:
        return "NVARCHAR(4000) NULL"
    return "NVARCHAR(MAX) NULL"


def _infer_sql_type(series: pd.Series, col_name: str) -> str:
    s = series.dropna()
    if s.empty:
        return "NVARCHAR(256) NULL"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "NVARCHAR(64) NULL"
    if pd.api.types.is_bool_dtype(series):
        return "NVARCHAR(8) NULL"
    if pd.api.types.is_integer_dtype(series):
        if _name_suggests_decimal(col_name):
            return "DECIMAL(18, 6) NULL"
        if _name_suggests_code_text(col_name):
            return _nvarchar_tier_from_series(series)
        return "INT NULL"
    if pd.api.types.is_float_dtype(series):
        s2 = s.astype(float)
        if _name_suggests_code_text(col_name):
            return _nvarchar_tier_from_series(series)
        if np.all(np.isfinite(s2)) and np.allclose(s2, np.round(s2), rtol=0, atol=1e-9):
            if _name_suggests_decimal(col_name):
                return "DECIMAL(18, 6) NULL"
            return "INT NULL"
        return "DECIMAL(18, 6) NULL"
    if series.dtype == object:
        sample = s.head(200)
        if len(sample) > 0:
            first = sample.iloc[0]
            if isinstance(first, (dt.datetime, pd.Timestamp)):
                return "NVARCHAR(64) NULL"
        num = pd.to_numeric(s, errors="coerce")
        if num.notna().all() and len(num) == len(s):
            if _name_suggests_code_text(col_name):
                return _nvarchar_tier_from_series(series)
            if np.allclose(num.to_numpy(), np.round(num.to_numpy()), rtol=0, atol=1e-9):
                if _name_suggests_decimal(col_name):
                    return "DECIMAL(18, 6) NULL"
                return "INT NULL"
            return "DECIMAL(18, 6) NULL"
    return _nvarchar_tier_from_series(s)


def _sql_type_to_kind(sql_t: str) -> str:
    if sql_t.startswith("INT"):
        return "int"
    if sql_t.startswith("DECIMAL"):
        return "decimal"
    return "nvarchar"


def _build_specs(df: pd.DataFrame) -> Tuple[List[str], List[str], List[str]]:
    col_names: List[str] = []
    for i, c in enumerate(df.columns):
        name = str(c).strip() if c is not None else ""
        if name == "":
            name = f"Column{i + 1}"
        col_names.append(name)
    df.columns = col_names
    sql_types = [_infer_sql_type(df[c], c) for c in col_names]
    kinds = [_sql_type_to_kind(t) for t in sql_types]
    return col_names, sql_types, kinds


def _build_create_ddl(table: str, col_names: Sequence[str], sql_types: Sequence[str]) -> str:
    lines = [f"CREATE TABLE dbo.{table} ("]
    for n, t in zip(col_names, sql_types):
        lines.append(f"    {_bracket_ident(n)} {t},")
    lines[-1] = lines[-1].rstrip(",")
    lines.append(");")
    return "\n".join(lines)


def _connect(server: str, database: str, user: str, password: str):
    import pyodbc

    drivers = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "SQL Server",
    ]
    last_err: Exception | None = None
    for drv in drivers:
        conn_s = (
            f"DRIVER={{{drv}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={user};"
            f"PWD={password};"
            "TrustServerCertificate=yes;"
        )
        try:
            return pyodbc.connect(conn_s, timeout=15)
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
    raise RuntimeError(f"Could not connect with any ODBC driver. Last error: {last_err}") from last_err


def _load_one_table(
    conn: Any,
    table: str,
    path: str,
    *,
    dry_run: bool,
) -> int:
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    df = pd.read_excel(path, sheet_name=0)
    col_names, sql_types, kinds = _build_specs(df)
    drop_sql = f"IF OBJECT_ID(N'dbo.{table}', N'U') IS NOT NULL DROP TABLE dbo.{table};"
    create_sql = _build_create_ddl(table, col_names, sql_types)
    if dry_run:
        print(f"-- {table} ({len(col_names)} columns, {len(df)} rows)")
        print(drop_sql)
        print(create_sql)
        print()
        return len(df)

    cur = conn.cursor()
    cur.execute(drop_sql)
    cur.execute(create_sql)
    bracketed = ", ".join(_bracket_ident(n) for n in col_names)
    placeholders = ", ".join("?" * len(col_names))
    insert_sql = f"INSERT INTO dbo.{table} ({bracketed}) VALUES ({placeholders})"
    batch: list[tuple[Any, ...]] = []
    for _, row in df.iterrows():
        raw = [row[c] for c in col_names]
        batch.append(tuple(_coerce(raw[i], kinds[i]) for i in range(len(col_names))))
    cur.fast_executemany = True
    cur.executemany(insert_sql, batch)
    conn.commit()
    return len(batch)


def main() -> int:
    p = argparse.ArgumentParser(
        description="Drop/recreate dbo.OITW_TMP, OITM_TMP, ORSG_TMP, BEG_INV, SHIP_SCH, … from Excel."
    )
    p.add_argument("--server", default=os.environ.get("SQLSERVER_SERVER", ""))
    p.add_argument("--database", default=os.environ.get("SQLSERVER_DATABASE", "SW"))
    p.add_argument("--user", default=os.environ.get("SQLSERVER_USER", ""))
    p.add_argument(
        "--oitw",
        default=os.environ.get(
            "OITW_TMP_XLSX",
            r"c:\Users\lenovo\Downloads\OITW_TMP.xlsx",
        ),
        help="Path to OITW_TMP.xlsx (skip with empty string)",
    )
    p.add_argument(
        "--oitm",
        default=os.environ.get(
            "OITM_TMP_XLSX",
            r"c:\Users\lenovo\Downloads\OITM_TMP.xlsx",
        ),
        help="Path to OITM_TMP.xlsx (skip with empty string)",
    )
    p.add_argument(
        "--orsg",
        default=os.environ.get(
            "ORSG_TMP_XLSX",
            r"c:\Users\lenovo\Downloads\ORSG_TMP.xlsx",
        ),
        help="Path to ORSG_TMP.xlsx (skip with empty string)",
    )
    p.add_argument(
        "--beg-inv",
        dest="beg_inv",
        default=os.environ.get(
            "BEG_INV_XLSX",
            r"c:\Users\lenovo\Downloads\BEG_INV.xlsx",
        ),
        help="Path to BEG_INV.xlsx (skip with empty string)",
    )
    p.add_argument(
        "--ship-sch",
        dest="ship_sch",
        default=os.environ.get(
            "SHIP_SCH_XLSX",
            r"c:\Users\lenovo\Downloads\SHIP_SCH.xlsx",
        ),
        help="Path to SHIP_SCH.xlsx (skip with empty string)",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Print DROP/CREATE DDL and row counts only; do not connect.",
    )
    p.add_argument("--no-oitw", action="store_true", help="Do not load dbo.OITW_TMP.")
    p.add_argument("--no-oitm", action="store_true", help="Do not load dbo.OITM_TMP.")
    p.add_argument("--no-orsg", action="store_true", help="Do not load dbo.ORSG_TMP.")
    p.add_argument(
        "--no-beg-inv",
        dest="no_beg_inv",
        action="store_true",
        help="Do not load dbo.BEG_INV.",
    )
    p.add_argument(
        "--no-ship-sch",
        dest="no_ship_sch",
        action="store_true",
        help="Do not load dbo.SHIP_SCH.",
    )
    args = p.parse_args()

    tasks: list[tuple[str, str]] = []
    if not args.no_oitw and args.oitw and str(args.oitw).strip():
        tasks.append(("OITW_TMP", args.oitw.strip()))
    if not args.no_oitm and args.oitm and str(args.oitm).strip():
        tasks.append(("OITM_TMP", args.oitm.strip()))
    if not args.no_orsg and args.orsg and str(args.orsg).strip():
        tasks.append(("ORSG_TMP", args.orsg.strip()))
    if not args.no_beg_inv and args.beg_inv and str(args.beg_inv).strip():
        tasks.append(("BEG_INV", args.beg_inv.strip()))
    if not args.no_ship_sch and args.ship_sch and str(args.ship_sch).strip():
        tasks.append(("SHIP_SCH", args.ship_sch.strip()))
    if not tasks:
        print(
            "Nothing to load: set --oitw / --oitm / --orsg / --beg-inv / --ship-sch paths or unset --no-* flags.",
            file=sys.stderr,
        )
        return 1

    password = os.environ.get("SQLSERVER_PASSWORD", "")
    if not args.dry_run and (not args.server or not args.user or not password):
        print(
            "Set SQLSERVER_PASSWORD and pass --server / --user (or use --dry-run).",
            file=sys.stderr,
        )
        return 1

    conn = None
    if not args.dry_run:
        conn = _connect(args.server, args.database, args.user, password)
    try:
        for table, path in tasks:
            n = _load_one_table(conn, table, path, dry_run=args.dry_run)
            if args.dry_run:
                print(f"Would load {n} row(s) into dbo.{table} from {path}")
            else:
                print(f"OK: dbo.{table} recreated and loaded {n} row(s) from {path}.")
        return 0
    finally:
        if conn is not None:
            conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
