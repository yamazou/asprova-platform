"""SAP Business One (SQL Server) 用 ERP アダプタ。

モジュール:
    - ``service``           : Bridge 向け ``SapB1BridgeService`` 実装
    - ``connection``        : SQL Server (pyodbc) 接続ヘルパ
    - ``schema``            : SAP B1 ステージングテーブルの列名解決ユーティリティ
    - ``integrated_master`` : ITT1_TMP / OITT_TMP からの統合マスタ取得
    - ``item_table``        : OITM_TMP からの品目テーブル取得
    - ``inventory_table``   : BEG_INV からの在庫テーブル取得
    - ``order_table``       : SHIP_SCH からの注文テーブル取得
"""
