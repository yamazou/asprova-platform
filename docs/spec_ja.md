# 仕様書（Asprova Platform 統合）

## 1. 開発環境
- 対象環境: Windows 10/11（例: 10.0.19045）
- 実行言語: Python
- Webフレームワーク: Flask
- 使用ライブラリ
  - viewer: `flask`, `sqlite3`, `csv`, `openpyxl`（Excel出力時）
  - bridge: `flask`, `oracledb`（Oracle接続・取得時）
- ポート
  - viewer: `5000`
  - bridge: `5001`

## 2. フォルダ構成
- `common/`
  - `templates/base.html`: 共通レイアウト（viewer/bridgeで共通利用）
  - `static/css/style.css`: 共通CSS
- `core/`
  - `csv_loader.py`: 共通CSV処理（区切り文字検出、DictReader生成、DB→CSV出力）
  - `asprova_parser.py`: Asprova形式CSVの列推定・1行パース（DB投入用）
- `apps/`
  - `viewer/`: スケジュール閲覧・CSV取り込み
  - `bridge/`: Oracle DB接続・CSVダウンロード
- `config/`: 必要に応じて設定ファイル
- `run.py`: viewer/bridge の起動

## 3. 起動方法
- 両方起動（デフォルト）
  - `py run.py`
- viewerのみ起動
  - `py run.py viewer`
- bridgeのみ起動
  - `py run.py bridge`
- 起動URL
  - viewer: `http://localhost:5000`
  - bridge: `http://localhost:5001`

## 4. 画面構成
共通UIは `common/templates/base.html` が提供し、各アプリは以下の画面を持ちます。

- 共通（ヘッダー/フッター）
  - viewer: Gantt / PSI / Import CSV へのナビ、Clearボタン（データがある場合）
  - bridge: Connectボタンと接続ステータス表示（モーダル経由でOracle接続）
- viewer
  - Schedule（スケジュール表示）
  - Gantt Chart（操作バー、ツールチップ、次工程リンク）
  - PSI Viewer（月次のSupply/Demand/Stock、Excel出力）
  - Import CSV（CSVアップロード）
- bridge
  - Index（各CSVダウンロード種別の選択カード）
  - Confirm/Connect（ダウンロード確認・Oracle接続モーダル）

## 5. 機能
- viewer
  - CSVアップロード
    - UTF-8 BOM対応デコード
    - 区切り文字（`,`/`\t`/`;`/`|`）自動検出
    - ヘッダから列マッピングを自動推定
  - SQLiteへ取り込み
    - schedule.db にスケジュール行を格納
    - 既存テーブルのカラム不足があればスキーマを更新
  - スケジュール表示
    - 週/2週/3週/月の期間切替
    - マシンでの絞り込み
  - Gantt表示
    - ツールチップ表示
    - 次工程リンク描画（絞り込み条件あり）
  - PSI表示
    - 月次のSupply/Demand/Stockテーブル表示
    - Excel出力
  - 全データ削除（Clear）
- bridge
  - Oracle DB接続（セッションに保存）
  - CSVダウンロード
    - Integrated Master
    - Item Table
    - Order Table
    - Resource Table
    - Inventory Table
  - ブラウザ上で保存先選択（ファイル保存UI）

## ファイルフォーマットのおすすめ
- ドキュメント: Markdown（`.md`）
  - `docs/spec_ja.md` のように管理しやすい
  - 差分が見やすく、レビューしやすい

