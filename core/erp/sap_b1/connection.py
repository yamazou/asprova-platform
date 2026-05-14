"""SQL Server ODBC connection (Bridge, sync scripts)."""

from __future__ import annotations

from typing import Any


def connect_sqlserver(
    server: str,
    database: str,
    user: str,
    password: str,
    *,
    timeout: int = 15,
) -> Any:
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
            f"TrustServerCertificate=yes;"
        )
        try:
            return pyodbc.connect(conn_s, timeout=timeout)
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
    raise RuntimeError(
        f"Could not connect to SQL Server. Check the ODBC driver. Last error: {last_err}"
    ) from last_err
