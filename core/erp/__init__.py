"""ERP 別のアダプタ。

サブパッケージ:
    - ``mcframe``: mcframe (Oracle) 用の取込ロジック
    - ``sap_b1``:  SAP Business One (SQL Server) 用の取込ロジック
    - ``excel``:   Excel ベースの ERP/手運用ファイル取込ロジック

顧客毎の納品時には、その顧客が利用する ERP サブパッケージのみを同梱する
運用を想定している (例: PHC 納品 = ``mcframe`` のみ、PEB 納品 = ``excel`` のみ)。
"""
