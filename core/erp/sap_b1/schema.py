"""SAP B1 ステージングテーブルの列名解決ユーティリティ。

SAP B1 の Excel エクスポート列名はスペース有無・大文字小文字が混在する
(例: ``Parent Item`` / ``ParentItem`` / ``parent_item``) ため、
論理名 → 実列名のゆるいマッチングを担う共通処理を集約している。

このモジュール内のヘルパは ``core/erp/sap_b1`` 配下の各テーブル取得
モジュールから利用される。
"""

from __future__ import annotations

from typing import Dict, List, Optional, Sequence


def norm_col_key(name: str) -> str:
    """``'Parent Item'`` / ``'ParentItem'`` / ``'parent_item'`` を等価視する。

    英数字以外を除去し、小文字化した文字列を返す。
    """

    return "".join(ch.lower() for ch in name if ch.isalnum())


def bracket_ident(col: str) -> str:
    """SQL Server 用に ``[列名]`` 形式に括る。``]`` を含む場合はエスケープ。"""

    return "[" + col.replace("]", "]]") + "]"


def sqlserver_table_columns(
    pyodbc_conn, table: str, schema: str = "dbo"
) -> List[str]:
    """``INFORMATION_SCHEMA.COLUMNS`` から実列名を取得する。"""

    cur = pyodbc_conn.cursor()
    cur.execute(
        """
        SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """,
        (schema, table),
    )
    return [row[0] for row in cur.fetchall()]


def pick_column(
    available: Sequence[str], candidates: Sequence[str]
) -> Optional[str]:
    """論理名候補リストから、実列名側にある最初の一致を返す。"""

    by_norm: Dict[str, str] = {}
    for c in available:
        n = norm_col_key(c)
        if n not in by_norm:
            by_norm[n] = c
    for cand in candidates:
        n = norm_col_key(cand)
        if n in by_norm:
            return by_norm[n]
    return None
