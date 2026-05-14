"""
Load SAP B1 dbo.ITT1_TMP + dbo.OITT_TMP, build rows like integrated_master.csv
(P_ITM_CD … PRODUCTION), and replace sqlite integrates (schedule.db).

Optional env:
  SAP_INTEGRATED_U_ROWS=0   — only I rows (no U/M line; default is one U/M per parent: PROD / 1D)

PowerShell example:

  $env:SQLSERVER_PASSWORD = '***'
  py -3 scripts/sync_integrated_master_to_schedule_db.py ^
    --server "LAPTOP-4ST122V3\\SQLEXPRESS" --database SW --user sa
"""

from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from config.settings import DB_PATH  # noqa: E402
from core.erp.sap_b1.integrated_master import sync_from_sqlserver_to_sqlite  # noqa: E402


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--server", default=os.environ.get("SQLSERVER_SERVER", ""))
    p.add_argument("--database", default=os.environ.get("SQLSERVER_DATABASE", "SW"))
    p.add_argument("--user", default=os.environ.get("SQLSERVER_USER", ""))
    p.add_argument("--sqlite", default=DB_PATH, help="Path to schedule.db")
    args = p.parse_args()

    pwd = os.environ.get("SQLSERVER_PASSWORD", "")
    if not args.server or not args.user or not pwd:
        print(
            "Set SQLSERVER_PASSWORD and --server / --user (or SQLSERVER_* env).",
            file=sys.stderr,
        )
        return 1

    os.makedirs(os.path.dirname(os.path.abspath(args.sqlite)), exist_ok=True)
    n = sync_from_sqlserver_to_sqlite(
        server=args.server,
        database=args.database,
        user=args.user,
        password=pwd,
        sqlite_db_path=args.sqlite,
    )
    print(f"OK: integrates table replaced with {n} row(s) at {args.sqlite}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
