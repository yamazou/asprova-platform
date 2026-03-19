import csv
import io
import os
import re
from datetime import datetime
import sys
from pathlib import Path
from jinja2 import ChoiceLoader, FileSystemLoader

from flask import Flask, Response, flash, redirect, render_template, request, send_file, session, url_for
import oracledb

# 古いバージョンの Oracle DB に接続するため、
# Python-oracledb を thick モード（Oracle Client 経由）で使用する。
try:
    # Windows では、Oracle Client / Instant Client が PATH やレジストリに
    # 登録されていれば、引数なしで自動検出されます。
    oracledb.init_oracle_client()
except oracledb.ProgrammingError:
    # すでに初期化済みの場合などは無視
    pass

PLATFORM_ROOT = Path(__file__).resolve().parents[2]
COMMON_DIR = PLATFORM_ROOT / "common"

sys.path.insert(0, str(PLATFORM_ROOT))
from core.csv_loader import rows_to_csv

app = Flask(__name__, static_folder=str(COMMON_DIR / "static"))
app.jinja_loader = ChoiceLoader(
    [FileSystemLoader(str(COMMON_DIR / "templates")), app.jinja_loader]
)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "asprova-bridge-dev-secret")

MCFRAME_LOGO_PATH = os.environ.get(
    "MCFRAME_LOGO_PATH",
    r"C:\Users\lenovo\.cursor\projects\c-Users-lenovo-asprova-bridge\assets\c__Users_lenovo_AppData_Roaming_Cursor_User_workspaceStorage_bc2da2fe61af596e292be7a2910fd7a2_images_HANA-FIRST-logo_ori-2d4ded98-70e0-42b6-ad82-dadb87d98261.png",
)

CONFIRM_LOGO_PATH = os.environ.get(
    "CONFIRM_LOGO_PATH",
    r"C:\Users\lenovo\.cursor\projects\c-Users-lenovo-asprova-bridge\assets\c__Users_lenovo_AppData_Roaming_Cursor_User_workspaceStorage_bc2da2fe61af596e292be7a2910fd7a2_images_mcframe-logo-b8b12ae0-3865-415a-9a41-0a5440ffc293.png",
)


HTML_INDEX = None


MASTER_SQL = """
SELECT
    b.P_ITM_CD AS P_ITM_CD,
    10 AS PROCESS_NO,
    '10' AS PROCESS_CD,
    'I' AS INST_TYP,
    'In' || ROW_NUMBER() OVER (
      PARTITION BY b.P_ITM_CD
      ORDER BY b.C_ITM_CD
    ) AS INST_CD,
    b.C_ITM_CD AS ITM_RESOURCE,
    TO_CHAR(b.C_REQ_QTY) AS PRODUCTION
FROM
    {schema}.SM_BOM_ALL b
WHERE
    b.BOM_PTN = 1

UNION ALL

SELECT
    hl.ITM_CD AS P_ITM_CD,
    10 AS PROCESS_NO,
    '10' AS PROCESS_CD,
    'U' AS INST_TYP,
    'M' AS INST_CD,
    hl.LINE_CD AS ITM_RESOURCE,
    TO_CHAR(hl.STD_LD) || 'sp' AS PRODUCTION
FROM
    {schema}.SM_HINLINE_ALL hl
WHERE
    hl.CO_CD = 'J0001'
"""


INTEGRATED_HEADERS = [
    "P_ITM_CD",
    "PROCESS_NO",
    "PROCESS_CD",
    "INST_TYP",
    "INST_CD",
    "ITM_RESOURCE",
    "PRODUCTION",
]


ITEM_TABLE_SQL = """
SELECT
    c.ITM_CD,
    c.ITM_NM,
    MAX(m.MAX_LOT_UNIT_QTY) AS MAX_LOT_UNIT_QTY
FROM
    {schema}.CM_HINMO_ALL c
    JOIN {schema}.SM_HINMOS_ALL m
      ON m.ITM_CD = c.ITM_CD
WHERE
    c.CO_CD = 'J0001'
    AND m.CO_CD = 'J0001'
    -- AND m.BOM_PTN = 1
GROUP BY
    c.ITM_CD,
    c.ITM_NM
ORDER BY
    c.ITM_CD
"""

ITEM_TABLE_HEADERS = [
    "ITM_CD",
    "ITM_NM",
    "MAX_LOT_UNIT_QTY",
]

ORDER_TABLE_SQL = """
SELECT
    REQ_NO,
    ITM_CD,
    DLV_DT,
    REQ_QTY
FROM
    {schema}.ST_SHOYO_ALL
ORDER BY
    REQ_NO,
    ITM_CD,
    DLV_DT
"""

ORDER_TABLE_HEADERS = [
    "REQ_NO",
    "ITM_CD",
    "DLV_DT",
    "REQ_QTY",
]

RESOURCE_TABLE_SQL = """
SELECT
    LINE_CD,
    LINE_NM
FROM
    {schema}.SM_LINE_ALL
ORDER BY
    LINE_CD
"""

RESOURCE_TABLE_HEADERS = [
    "LINE_CD",
    "LINE_NM",
]

INVENTORY_TABLE_SQL = """
SELECT
    'INV' || LPAD(ROW_NUMBER() OVER (ORDER BY ITM_CD), 5, '0') AS INV_CD,
    ITM_CD,
    STK_QTY,
    CAST(NULL AS VARCHAR2(10)) AS INV_DT
FROM
    {schema}.ST_GNZAIKO_ALL
ORDER BY
    ITM_CD
"""

INVENTORY_TABLE_HEADERS = [
    "INV_CD",
    "ITM_CD",
    "STK_QTY",
    "INV_DT",
]


def get_connection():
    user = session.get("oracle_user")
    password = session.get("oracle_password")
    if not (user and password):
        raise RuntimeError("未接続です。先に「Connect mcframe」からOracleへ接続してください。")
    dsn = "orcl"
    return oracledb.connect(user=user, password=password, dsn=dsn)


def get_schema() -> str:
    schema = session.get("oracle_schema")
    if not schema:
        raise RuntimeError("未接続です。先に「Connect mcframe」からOracleへ接続してください。")
    return schema


_SCHEMA_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,29}$")


def fetch_rows(sql: str):
    with get_connection() as conn:
        with conn.cursor() as cur:
            schema = get_schema()
            cur.execute(sql.format(schema=schema))
            for row in cur:
                yield row


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "bridge_index.html", oracle_connected=bool(session.get("oracle_connected"))
    )


@app.route("/assets/mcframe-logo.png", methods=["GET"])
def mcframe_logo():
    return send_file(MCFRAME_LOGO_PATH, mimetype="image/png")


@app.route("/assets/confirm-logo.png", methods=["GET"])
def confirm_logo():
    return send_file(CONFIRM_LOGO_PATH, mimetype="image/png")


@app.route("/connect", methods=["POST"])
def connect_oracle():
    oracle_id = request.form.get("oracle_id", "").strip()
    oracle_pwd = request.form.get("oracle_pwd", "")
    oracle_schema = request.form.get("oracle_schema", "").strip()
    if not oracle_id or not oracle_pwd or not oracle_schema:
        flash("ID / Password / Schema を入力してください。", "error")
        return redirect(url_for("index"))

    if not _SCHEMA_RE.match(oracle_schema):
        flash("Schema は英数字とアンダースコアのみ（先頭は英字、最大30文字）で入力してください。", "error")
        return redirect(url_for("index"))

    oracle_schema = oracle_schema.upper()

    # ここで初めてOracleに接続し、成功したらセッションに保持
    try:
        conn = oracledb.connect(user=oracle_id, password=oracle_pwd, dsn="orcl")
        conn.close()
    except Exception as exc:  # noqa: BLE001
        session.pop("oracle_user", None)
        session.pop("oracle_password", None)
        session.pop("oracle_schema", None)
        session.pop("oracle_connected", None)
        flash(f"接続に失敗しました: {exc}", "error")
        return redirect(url_for("index"))

    session["oracle_user"] = oracle_id
    session["oracle_password"] = oracle_pwd
    session["oracle_schema"] = oracle_schema
    session["oracle_connected"] = True
    flash("接続に成功しました。", "success")
    return redirect(url_for("index"))


@app.route("/download/integrated", methods=["POST"])
def download_integrated():
    try:
        rows = list(fetch_rows(MASTER_SQL))
        csv_data = rows_to_csv(rows, INTEGRATED_HEADERS)
    except Exception as exc:  # noqa: BLE001
        flash(f"エラーが発生しました: {exc}", "error")
        return redirect(url_for("index"))

    filename = "integrated_master.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/download/item-table", methods=["POST"])
def download_item_table():
    try:
        rows = list(fetch_rows(ITEM_TABLE_SQL))
        csv_data = rows_to_csv(rows, ITEM_TABLE_HEADERS)
    except Exception as exc:  # noqa: BLE001
        flash(f"エラーが発生しました: {exc}", "error")
        return redirect(url_for("index"))

    filename = "item_table.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/download/order-table", methods=["POST"])
def download_order_table():
    try:
        rows = list(fetch_rows(ORDER_TABLE_SQL))
        csv_data = rows_to_csv(rows, ORDER_TABLE_HEADERS)
    except Exception as exc:  # noqa: BLE001
        flash(f"エラーが発生しました: {exc}", "error")
        return redirect(url_for("index"))

    filename = "order_table.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/download/resource-table", methods=["POST"])
def download_resource_table():
    try:
        rows = list(fetch_rows(RESOURCE_TABLE_SQL))
        csv_data = rows_to_csv(rows, RESOURCE_TABLE_HEADERS)
    except Exception as exc:  # noqa: BLE001
        flash(f"エラーが発生しました: {exc}", "error")
        return redirect(url_for("index"))

    filename = "resource_table.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/download/inventory-table", methods=["POST"])
def download_inventory_table():
    try:
        rows = list(fetch_rows(INVENTORY_TABLE_SQL))
        csv_data = rows_to_csv(rows, INVENTORY_TABLE_HEADERS)
    except Exception as exc:  # noqa: BLE001
        flash(f"エラーが発生しました: {exc}", "error")
        return redirect(url_for("index"))

    filename = "inventory_table.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5001")), debug=True)

