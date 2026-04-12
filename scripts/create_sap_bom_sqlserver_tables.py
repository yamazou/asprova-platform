"""
Create dbo.OITT_TMP and dbo.ITT1_TMP on SQL Server (SAP B1 BOM export layout).

Do not commit credentials. Example (PowerShell):

  $env:SQLSERVER_PASSWORD = 'your-password'
  py -3 scripts/create_sap_bom_sqlserver_tables.py ^
    --server "LAPTOP-4ST122V3\\SQLEXPRESS" --database SW --user sa
"""

from __future__ import annotations

import argparse
import os
import sys


DDL_OITT = """
IF OBJECT_ID(N'dbo.OITT_TMP', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OITT_TMP (
        [#]                              INT            NULL,
        [Parent Item]                    NVARCHAR(64)   NULL,
        [BOM Type]                       NVARCHAR(8)    NULL,
        [Price List]                     INT            NULL,
        [No. of Units]                   INT            NULL,
        [Creation Date]                  NVARCHAR(32)   NULL,
        [Date of Update]                 NVARCHAR(32)   NULL,
        [Postponed to Next Year]         NVARCHAR(8)    NULL,
        [Data source]                    NVARCHAR(8)    NULL,
        [User Signature]                 INT            NULL,
        [SCN Counter]                    INT            NULL,
        [Display Currency]               INT            NULL,
        [Whse for Finished Product]      NVARCHAR(16)   NULL,
        [Object Type]                    NVARCHAR(16)   NULL,
        [Log Instance - History]         INT            NULL,
        [Updating User]                  INT            NULL,
        [Distribution Rule]              NVARCHAR(128)  NULL,
        [Hide Components in Printing]    NVARCHAR(8)    NULL,
        [Distribution Rule2]             NVARCHAR(128)  NULL,
        [Distribution Rule3]             NVARCHAR(128)  NULL,
        [Distribution Rule4]             NVARCHAR(128)  NULL,
        [Distribution Rule5]             NVARCHAR(128)  NULL,
        [Time of Update]                 NVARCHAR(32)   NULL,
        [Project Code]                   NVARCHAR(64)   NULL,
        [Planned Average Production Size] INT           NULL,
        [Product Description]            NVARCHAR(512)  NULL,
        [Create Time - Incl. Secs]       BIGINT         NULL,
        [Update Full Time]               BIGINT         NULL,
        [Attachment Entry]               INT            NULL,
        [Attachments]                    INT            NULL
    );
END
"""

DDL_ITT1 = """
IF OBJECT_ID(N'dbo.ITT1_TMP', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ITT1_TMP (
        [#]                          INT             NULL,
        [Parent Item]                NVARCHAR(64)    NULL,
        [Component Element Number]   INT             NULL,
        [Visual Order]               INT             NULL,
        [Component Code]           NVARCHAR(64)    NULL,
        [Quantity]                   DECIMAL(18, 6)  NULL,
        [Warehouse]                  NVARCHAR(16)    NULL,
        [Price]                      DECIMAL(18, 6)  NULL,
        [Currency]                   NVARCHAR(16)    NULL,
        [Price List]                 INT             NULL,
        [Original Price]             DECIMAL(18, 6)  NULL,
        [Original Currency]          NVARCHAR(16)    NULL,
        [Issue Method]               NVARCHAR(8)     NULL,
        [Inventory UoM]              NVARCHAR(32)    NULL,
        [Comment]                    NVARCHAR(512)   NULL,
        [Log Instance]               INT             NULL,
        [Object]                     NVARCHAR(16)    NULL,
        [Distribution Rule]          NVARCHAR(128)   NULL,
        [Distribution Rule2]         NVARCHAR(128)   NULL,
        [Distribution Rule3]         NVARCHAR(128)   NULL,
        [Distribution Rule4]         NVARCHAR(128)   NULL,
        [Distribution Rule5]         NVARCHAR(128)   NULL,
        [Principal Input]            NVARCHAR(8)     NULL,
        [Project Code]               NVARCHAR(64)    NULL,
        [Component Type]             INT             NULL,
        [WIP Account Code]           NVARCHAR(64)    NULL,
        [Additional Quantity]        DECIMAL(18, 6)  NULL,
        [Row Text]                   NVARCHAR(512)   NULL,
        [Stage ID]                   NVARCHAR(64)    NULL,
        [Item Description]           NVARCHAR(512)   NULL
    );
END
"""


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


def main() -> int:
    p = argparse.ArgumentParser(description="Create SAP B1 BOM staging tables on SQL Server.")
    p.add_argument("--server", default=os.environ.get("SQLSERVER_SERVER", ""), help="e.g. HOST\\SQLEXPRESS")
    p.add_argument("--database", default=os.environ.get("SQLSERVER_DATABASE", "SW"))
    p.add_argument("--user", default=os.environ.get("SQLSERVER_USER", ""))
    args = p.parse_args()

    password = os.environ.get("SQLSERVER_PASSWORD", "")
    if not args.server or not args.user or not password:
        print(
            "Set SQLSERVER_PASSWORD and pass --server / --user (or SQLSERVER_* env vars).",
            file=sys.stderr,
        )
        return 1

    conn = _connect(args.server, args.database, args.user, password)
    try:
        cur = conn.cursor()
        cur.execute(DDL_OITT)
        cur.execute(DDL_ITT1)
        conn.commit()
        print("OK: dbo.OITT_TMP and dbo.ITT1_TMP are present (created if missing).")
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
