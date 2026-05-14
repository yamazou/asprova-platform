"""ERP 非依存のパーサ・ローダ。

ここに置くものの基準:
    - 特定 ERP / 顧客に依存しない汎用処理 (CSV / Asprova の PSI 取込など)
    - 出力は Asprova / 自前 SQLite の中立な表現

ERP 接続や SAP B1 / mcframe / Excel 固有のスキーマ処理は ``core.erp.<system>``
に置くこと。顧客固有の表示・集計差分は ``core.customers.<id>`` に置くこと。
"""
