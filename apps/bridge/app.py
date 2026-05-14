import csv
import io
import mimetypes
import os
import re
import sqlite3
from datetime import timedelta
import sys
from pathlib import Path
from jinja2 import ChoiceLoader, FileSystemLoader

from flask import (
    Flask,
    Response,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.utils import secure_filename


# oracledb は mcframe (Oracle) 顧客向け納品でしか必要ない。
# 納品時に oracledb / Oracle Client を含めない構成 (例: PEB のみ Excel 納品)
# でもアプリ起動できるよう、関数内で必要時に lazy import する。


def _ensure_oracledb_initialized() -> None:
    """mcframe 接続の前に Oracle Client を初期化する (idempotent)。"""

    try:
        import oracledb  # noqa: PLC0415

        try:
            oracledb.init_oracle_client()
        except oracledb.ProgrammingError:
            # すでに初期化済みの場合などは無視
            pass
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "oracledb is required for mcframe (Oracle) connections: "
            "run pip install oracledb."
        ) from exc

PLATFORM_ROOT = Path(__file__).resolve().parents[2]
COMMON_DIR = PLATFORM_ROOT / "common"

sys.path.insert(0, str(PLATFORM_ROOT))
from core.parsers.csv_loader import rows_to_csv
from core.erp._base import BridgeErpService, NotSupportedError
from core.customers import get_customer
from config.bridge_customers import BRIDGE_CUSTOMERS
from config.settings import DB_PATH

app = Flask(__name__, static_folder=str(COMMON_DIR / "static"))
app.jinja_loader = ChoiceLoader(
    [FileSystemLoader(str(COMMON_DIR / "templates")), app.jinja_loader]
)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "asprova-bridge-dev-secret")


@app.context_processor
def _inject_customer_view():
    """テンプレート全体で使えるよう ``customer_view`` を注入する。

    顧客 ID 直書きの ``{% if selected_customer_id == 'xxx' %}`` を
    ``{% for btn in customer_view.bridge_buttons %}`` のような
    顧客非依存の表現に置き換えるため。
    """

    return {
        "customer_view": get_customer(session.get("customer_id")).to_view(),
    }
# ブラウザを閉じても接続フォーム用セッションを残す（デフォルトのセッション Cookie は終了時に消える）
_bridge_session_days = int(os.environ.get("BRIDGE_SESSION_DAYS", "30"))
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=max(1, min(_bridge_session_days, 365)))

MCFRAME_LOGO_PATH = os.environ.get(
    "MCFRAME_LOGO_PATH",
    str(COMMON_DIR / "static" / "mcframe-logo.webp"),
)

CONFIRM_LOGO_PATH = os.environ.get(
    "CONFIRM_LOGO_PATH",
    str(COMMON_DIR / "static" / "mcframe-logo.webp"),
)


HTML_INDEX = None


# CSV 出力ヘッダ (ERP 横断で同一)。
# SQL 文や Excel 解析ロジックは ``core/erp/<system>/service.py`` 側に集約。
INTEGRATED_HEADERS = [
    "P_ITM_CD",
    "PROCESS_NO",
    "PROCESS_CD",
    "INST_TYP",
    "INST_CD",
    "ITM_RESOURCE",
    "PRODUCTION",
]

ITEM_TABLE_HEADERS = [
    "ITM_CD",
    "ITM_NM",
    "ITM_TYP",
    "MAX_LOT_UNIT_QTY",
]

ORDER_TABLE_HEADERS = [
    "REQ_NO",
    "ITM_CD",
    "DLV_DT",
    "REQ_QTY",
    "CUST_CD",
]

RESOURCE_TABLE_HEADERS = [
    "LINE_CD",
    "LINE_NM",
    "RESOURCE_GRP",
    "Sort_Order",
]

INVENTORY_TABLE_HEADERS = [
    "INV_CD",
    "ITM_CD",
    "STK_QTY",
    "INV_DT",
]

RESOURCE_LINE_CODE_CANDIDATES = ("LINE_CD", "Line Code", "LINE CODE")
RESOURCE_LINE_NAME_CANDIDATES = ("LINE_NM", "Line Name", "LINE NAME")
RESOURCE_CYCLE_TIME_CANDIDATES = ("CYCLE_TIME", "Cycle Time", "CYCLE TIME")
RESOURCE_CYCLE_TIME_HEADER = "CYCLE_TIME"
RESOURCE_CYCLE_TABLE = "bridge_resource_cycle_times"
RESOURCE_CYCLE_SOURCE_SESSION_KEY = "resource_cycle_source_name"
RESOURCE_CYCLE_DB_PATH = Path(DB_PATH)
RESOURCE_CYCLE_LEGACY_CO_CD = (
    os.environ.get("BRIDGE_RESOURCE_CYCLE_LEGACY_CO_CD", "NCI").strip().upper() or "NCI"
)


def _pick_first_matching(headers: list[str], candidates: tuple[str, ...]) -> str | None:
    for key in candidates:
        if key in headers:
            return key
    return None


def _open_resource_cycle_db() -> sqlite3.Connection:
    RESOURCE_CYCLE_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(RESOURCE_CYCLE_DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def _resource_cycle_scope_key() -> str:
    """Line Cycle Time のデータ分離キー。mcframe は co_cd、excel は customer_id。"""
    erp = get_erp_system()
    customer_id = str(session.get("customer_id") or "").strip().lower()
    if erp == "excel" and customer_id:
        return customer_id
    raw = (session.get("mcframe_co_cd") or "").strip().upper()
    if raw and _MCFRAME_CO_CD_RE.match(raw):
        return raw
    if customer_id:
        return customer_id
    return RESOURCE_CYCLE_LEGACY_CO_CD


def _ensure_resource_cycle_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {RESOURCE_CYCLE_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            co_cd TEXT NOT NULL DEFAULT '',
            line_code TEXT NOT NULL,
            line_name TEXT NOT NULL DEFAULT '',
            cycle_time TEXT NOT NULL DEFAULT '',
            source_filename TEXT NOT NULL DEFAULT '',
            sort_order INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur = conn.execute(f"PRAGMA table_info({RESOURCE_CYCLE_TABLE})")
    existing_cols = {str(row["name"]) for row in cur.fetchall()}
    if "co_cd" not in existing_cols:
        conn.execute(
            f"ALTER TABLE {RESOURCE_CYCLE_TABLE} ADD COLUMN co_cd TEXT NOT NULL DEFAULT ''"
        )
        conn.execute(
            f"""
            UPDATE {RESOURCE_CYCLE_TABLE}
            SET co_cd = ?
            WHERE TRIM(COALESCE(co_cd, '')) = ''
            """,
            (RESOURCE_CYCLE_LEGACY_CO_CD,),
        )


def _fetch_resource_cycle_rows() -> list[dict]:
    co_cd = _resource_cycle_scope_key()
    with _open_resource_cycle_db() as conn:
        _ensure_resource_cycle_table(conn)
        cur = conn.execute(
            f"""
            SELECT id, line_code, line_name, cycle_time, sort_order
            FROM {RESOURCE_CYCLE_TABLE}
            WHERE co_cd = ?
            ORDER BY sort_order ASC, id ASC
            """,
            (co_cd,),
        )
        return [
            {
                "id": int(r["id"]),
                "line_code": str(r["line_code"] or "").strip(),
                "line_name": str(r["line_name"] or "").strip(),
                "cycle_time": str(r["cycle_time"] or "").strip(),
                "sort_order": int(r["sort_order"] or 0),
            }
            for r in cur.fetchall()
        ]


def _replace_resource_cycle_rows(rows: list[dict], source_filename: str, co_cd: str) -> None:
    with _open_resource_cycle_db() as conn:
        _ensure_resource_cycle_table(conn)
        conn.execute(f"DELETE FROM {RESOURCE_CYCLE_TABLE} WHERE co_cd = ?", (co_cd,))
        conn.executemany(
            f"""
            INSERT INTO {RESOURCE_CYCLE_TABLE}
            (co_cd, line_code, line_name, cycle_time, source_filename, sort_order, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            [
                (
                    co_cd,
                    str(row.get("line_code") or "").strip(),
                    str(row.get("line_name") or "").strip(),
                    str(row.get("cycle_time") or "").strip(),
                    source_filename,
                    idx,
                )
                for idx, row in enumerate(rows)
            ],
        )
        conn.commit()


def _update_resource_cycle_rows_by_id(updates: dict[int, dict[str, str | int]]) -> None:
    if not updates:
        return
    co_cd = _resource_cycle_scope_key()
    with _open_resource_cycle_db() as conn:
        _ensure_resource_cycle_table(conn)
        for row_id, payload in updates.items():
            cycle_time = str(payload.get("cycle_time") or "").strip()
            sort_order = int(payload.get("sort_order") or 0)
            conn.execute(
                f"""
                UPDATE {RESOURCE_CYCLE_TABLE}
                SET cycle_time = ?, sort_order = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ? AND co_cd = ?
                """,
                (cycle_time, sort_order, row_id, co_cd),
            )
        conn.commit()


def _delete_resource_cycle_row_by_id(row_id: int) -> int:
    co_cd = _resource_cycle_scope_key()
    with _open_resource_cycle_db() as conn:
        _ensure_resource_cycle_table(conn)
        cur = conn.execute(
            f"DELETE FROM {RESOURCE_CYCLE_TABLE} WHERE id = ? AND co_cd = ?",
            (row_id, co_cd),
        )
        conn.commit()
        return int(cur.rowcount or 0)


def _insert_resource_cycle_row(
    *,
    line_code: str,
    line_name: str,
    cycle_time: str,
    sort_order: int,
) -> int:
    """Cycle Time Master に手動で 1 行追加する。

    現在の scope key (mcframe は co_cd、excel は customer_id) 単位で重複チェックし、
    同じ ``line_code`` が既に存在する場合は ``RuntimeError`` を送出する。
    source_filename には ``(manual)`` を埋めて、CSV 取り込み行と区別できるようにする。
    """

    co_cd = _resource_cycle_scope_key()
    line_code = (line_code or "").strip()
    if not line_code:
        raise RuntimeError("Line Code is required.")
    line_name = (line_name or "").strip()
    cycle_time = (cycle_time or "").strip()
    with _open_resource_cycle_db() as conn:
        _ensure_resource_cycle_table(conn)
        cur = conn.execute(
            f"SELECT id FROM {RESOURCE_CYCLE_TABLE} "
            f"WHERE co_cd = ? AND line_code = ?",
            (co_cd, line_code),
        )
        if cur.fetchone():
            raise RuntimeError(
                f"The same Line Code already exists: {line_code}"
            )
        cur = conn.execute(
            f"""
            INSERT INTO {RESOURCE_CYCLE_TABLE}
            (co_cd, line_code, line_name, cycle_time, source_filename, sort_order, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (co_cd, line_code, line_name, cycle_time, "(manual)", sort_order),
        )
        conn.commit()
        return int(cur.lastrowid or 0)


def _load_cycle_time_map_from_sqlite() -> dict[str, str]:
    co_cd = _resource_cycle_scope_key()
    try:
        with _open_resource_cycle_db() as conn:
            _ensure_resource_cycle_table(conn)
            cur = conn.execute(
                f"""
                SELECT line_code, cycle_time
                FROM {RESOURCE_CYCLE_TABLE}
                WHERE co_cd = ?
                ORDER BY sort_order ASC, id ASC
                """,
                (co_cd,),
            )
            out: dict[str, str] = {}
            for row in cur.fetchall():
                line_code = str(row["line_code"] or "").strip()
                cycle_time = str(row["cycle_time"] or "").strip()
                if not line_code or not cycle_time:
                    continue
                out[line_code] = cycle_time
            return out
    except Exception:
        return {}


def _load_sort_order_map_from_sqlite() -> dict[str, int]:
    co_cd = _resource_cycle_scope_key()
    try:
        with _open_resource_cycle_db() as conn:
            _ensure_resource_cycle_table(conn)
            cur = conn.execute(
                f"""
                SELECT line_code, sort_order
                FROM {RESOURCE_CYCLE_TABLE}
                WHERE co_cd = ?
                ORDER BY sort_order ASC, id ASC
                """,
                (co_cd,),
            )
            out: dict[str, int] = {}
            for row in cur.fetchall():
                line_code = str(row["line_code"] or "").strip()
                if not line_code:
                    continue
                out[line_code] = int(row["sort_order"] or 0)
            return out
    except Exception:
        return {}


def _apply_cycle_time_to_integrated_records(records: list[dict]) -> list[dict]:
    cycle_map = _load_cycle_time_map_from_sqlite()
    if not cycle_map:
        return records
    for rec in records:
        itm_resource = str(rec.get("ITM_RESOURCE") or "").strip()
        if itm_resource and itm_resource in cycle_map:
            rec["PRODUCTION"] = cycle_map[itm_resource]
    return records


def _read_resource_cycle_rows_from_bytes(raw: bytes) -> list[dict]:
    errors: list[Exception] = []
    for enc in ("utf-8-sig", "cp932", "shift_jis"):
        try:
            text = raw.decode(enc)
            reader = csv.DictReader(io.StringIO(text))
            headers = list(reader.fieldnames or [])
            rows = [dict(r) for r in reader]
            line_key = _pick_first_matching(headers, RESOURCE_LINE_CODE_CANDIDATES)
            if not line_key:
                raise RuntimeError("The CSV does not contain a LINE_CD or Line Code column.")
            name_key = _pick_first_matching(headers, RESOURCE_LINE_NAME_CANDIDATES) or "LINE_NM"
            cycle_key = _pick_first_matching(headers, RESOURCE_CYCLE_TIME_CANDIDATES)
            if not cycle_key:
                cycle_key = RESOURCE_CYCLE_TIME_HEADER
            out: list[dict] = []
            for row in rows:
                out.append(
                    {
                        "line_code": str(row.get(line_key) or "").strip(),
                        "line_name": str(row.get(name_key) or "").strip(),
                        "cycle_time": str(row.get(cycle_key) or "").strip(),
                    }
                )
            return out
        except UnicodeDecodeError as exc:
            errors.append(exc)
            continue
    raise RuntimeError("Could not determine the CSV character encoding.") from (
        errors[-1] if errors else None
    )


def get_schema() -> str:
    schema = session.get("oracle_schema")
    if not schema:
        raise RuntimeError("Not connected. Please connect to the database from Connect first.")
    return schema


def _get_erp_service() -> BridgeErpService:
    """セッション情報から ERP 別サービスを生成して返す (lazy import)。

    ここでだけ ``core.erp.<system>.service`` を import することで、
    顧客納品時に未使用の ERP サブパッケージを物理削除しても
    アプリ起動時の ImportError を避けられる。
    """

    erp = get_erp_system()
    if erp == "mcframe":
        from core.erp.mcframe.service import McframeBridgeService  # noqa: PLC0415

        return McframeBridgeService(
            oracle_user=str(session.get("oracle_user") or ""),
            oracle_password=str(session.get("oracle_password") or ""),
            oracle_dsn=str(session.get("oracle_dsn") or ""),
            oracle_schema=get_schema(),
            mcframe_co_cd=get_mcframe_co_cd(),
        )
    if erp == "sap_b1":
        from core.erp.sap_b1.service import SapB1BridgeService  # noqa: PLC0415

        return SapB1BridgeService(
            server=str(session.get("oracle_dsn") or "").strip(),
            database=str(session.get("oracle_schema") or "").strip(),
            user=str(session.get("oracle_user") or ""),
            password=str(session.get("oracle_password") or ""),
        )
    if erp == "excel":
        from core.erp.excel.service import ExcelBridgeService  # noqa: PLC0415

        profile = _get_selected_customer_profile()
        if not profile:
            raise RuntimeError(
                "Please select a Customer for Excel import."
            )
        return ExcelBridgeService(
            customer=get_customer(session.get("customer_id")),
            profile=profile,
        )
    raise RuntimeError(f"Unsupported ERP type: {erp}")


def _verify_connection_alive() -> bool:
    """セッション上の ``oracle_connected`` が現在も実接続可能かを軽量検証する。

    フラグが立っていても DB が停止している場合は ``session["oracle_connected"]``
    を ``False`` に書き戻し、ヘッダ表示を ``DISCONNECTED`` に切り替える。
    Excel 顧客のように接続を伴わない ERP は常に True 扱い。
    """

    if not session.get("oracle_connected"):
        return False
    try:
        service = _get_erp_service()
    except Exception:
        # 接続情報が壊れている / 不足している → 切断扱い
        session["oracle_connected"] = False
        return False
    try:
        service.ping()
    except Exception:
        session["oracle_connected"] = False
        return False
    return True


_SCHEMA_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,29}$")
# SAP B1 / SQL Server: database name (SCHEMA 欄に DB 名を入力)
_SAP_B1_DB_RE = re.compile(r"^[A-Za-z0-9_.\-]{1,128}$")
# mcframe テーブルの CO_CD（Connect 画面の Company CD）。英数字 1〜20、保存時は大文字。
_MCFRAME_CO_CD_RE = re.compile(r"^[A-Z0-9]{1,20}$")


def get_mcframe_co_cd() -> str:
    """接続時にセッションへ保存した mcframe 用会社コード。未設定時は J0001。"""
    raw = (session.get("mcframe_co_cd") or "").strip().upper()
    if raw and _MCFRAME_CO_CD_RE.match(raw):
        return raw
    return "J0001"


def get_erp_system() -> str:
    raw = (session.get("erp_system") or "mcframe").strip().lower()
    return raw if raw in ("mcframe", "sap_b1", "excel") else "mcframe"


def _customer_profile_options() -> list[dict[str, str]]:
    options: list[dict[str, str]] = []
    for key, cfg in BRIDGE_CUSTOMERS.items():
        options.append(
            {
                "id": key,
                "label": str(cfg.get("label") or key),
                "erp_system": str(cfg.get("erp_system") or "mcframe"),
                "oracle_id": str(cfg.get("oracle_id") or ""),
                "oracle_pwd": str(cfg.get("oracle_pwd") or ""),
                "oracle_schema": str(cfg.get("oracle_schema") or ""),
                "oracle_dsn": str(cfg.get("oracle_dsn") or ""),
                "mcframe_co_cd": str(cfg.get("mcframe_co_cd") or ""),
                "excel_base_dir": str(cfg.get("excel_base_dir") or ""),
                "excel_integrated_file": str(cfg.get("excel_integrated_file") or ""),
                "excel_item_file": str(cfg.get("excel_item_file") or ""),
                "excel_order_file": str(cfg.get("excel_order_file") or ""),
                "excel_prd_plan_file": str(cfg.get("excel_prd_plan_file") or ""),
                "excel_resource_file": str(cfg.get("excel_resource_file") or ""),
                "excel_inventory_file": str(cfg.get("excel_inventory_file") or ""),
                "excel_inventory_wip_file": str(cfg.get("excel_inventory_wip_file") or ""),
            }
        )
    return options


def _get_selected_customer_profile() -> dict[str, str] | None:
    """セッションで選択中の顧客プロファイル (BRIDGE_CUSTOMERS のエントリ) を返す。"""

    customer_id = str(session.get("customer_id") or "").strip()
    if not customer_id:
        return None
    profile = BRIDGE_CUSTOMERS.get(customer_id)
    return profile if isinstance(profile, dict) else None


@app.route("/", methods=["GET"])
def index():
    show_line_cycle_master = (request.args.get("show_line_cycle") or "").strip() == "1"
    show_monthly_result = (request.args.get("show_monthly_result") or "").strip() == "1"
    line_cycle_rows: list[dict] = []
    if show_line_cycle_master:
        try:
            line_cycle_rows = _fetch_resource_cycle_rows()
        except Exception as exc:  # noqa: BLE001
            flash(f"Failed to load cycle time SQLite data: {exc}", "error")
    oracle_connected = _verify_connection_alive()
    return render_template(
        "bridge_index.html",
        oracle_connected=oracle_connected,
        erp_system=session.get("erp_system") or "mcframe",
        customer_profiles=_customer_profile_options(),
        selected_customer_id=str(session.get("customer_id") or ""),
        show_line_cycle_master=show_line_cycle_master,
        line_cycle_rows=line_cycle_rows,
        show_monthly_result=show_monthly_result,
        monthly_result_rows=[],
        monthly_result_source_name="",
    )


@app.route("/monthly-result", methods=["POST"])
def monthly_result():
    customer = get_customer(session.get("customer_id"))
    if not customer.supports_monthly_result():
        flash(
            "Monthly Result is not supported for this customer.",
            "error",
        )
        return redirect(url_for("index"))
    upload = request.files.get("monthly_result_excel")
    if not upload or not (upload.filename or "").strip():
        flash("Please select a Monthly Result source Excel file.", "error")
        return redirect(url_for("index", show_monthly_result="1"))
    try:
        raw = upload.read()
        if not raw:
            raise RuntimeError("The uploaded file is empty.")
        rows = customer.parse_monthly_result(raw)
        source_name = secure_filename(upload.filename or "monthly_result.xlsx")
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to load Monthly Result: {exc}", "error")
        return redirect(url_for("index", show_monthly_result="1"))

    show_line_cycle_master = (request.args.get("show_line_cycle") or "").strip() == "1"
    line_cycle_rows: list[dict] = []
    if show_line_cycle_master:
        try:
            line_cycle_rows = _fetch_resource_cycle_rows()
        except Exception as exc:  # noqa: BLE001
            flash(f"Failed to load cycle time SQLite data: {exc}", "error")
    oracle_connected = _verify_connection_alive()
    return render_template(
        "bridge_index.html",
        oracle_connected=oracle_connected,
        erp_system=session.get("erp_system") or "mcframe",
        customer_profiles=_customer_profile_options(),
        selected_customer_id=str(session.get("customer_id") or ""),
        show_line_cycle_master=show_line_cycle_master,
        line_cycle_rows=line_cycle_rows,
        show_monthly_result=True,
        monthly_result_rows=rows,
        monthly_result_source_name=source_name,
    )


def _send_logo_file(path: str) -> Response:
    mime, _ = mimetypes.guess_type(path)
    return send_file(path, mimetype=mime or "image/png")


@app.route("/assets/mcframe-logo.png", methods=["GET"])
def mcframe_logo():
    return _send_logo_file(MCFRAME_LOGO_PATH)


@app.route("/assets/confirm-logo.png", methods=["GET"])
def confirm_logo():
    return _send_logo_file(CONFIRM_LOGO_PATH)


@app.route("/connect", methods=["POST"])
def connect_oracle():
    customer_id = (request.form.get("customer_id") or "").strip()
    selected_profile = BRIDGE_CUSTOMERS.get(customer_id)
    oracle_id = request.form.get("oracle_id", "").strip()
    oracle_pwd = request.form.get("oracle_pwd", "")
    oracle_schema = request.form.get("oracle_schema", "").strip()
    oracle_dsn = request.form.get("oracle_dsn", "").strip()
    if selected_profile:
        if not oracle_id:
            oracle_id = str(selected_profile.get("oracle_id") or "").strip()
        if not oracle_pwd:
            oracle_pwd = str(selected_profile.get("oracle_pwd") or "")
        if not oracle_schema:
            oracle_schema = str(selected_profile.get("oracle_schema") or "").strip()
        if not oracle_dsn:
            oracle_dsn = str(selected_profile.get("oracle_dsn") or "").strip()
    erp_raw = (request.form.get("erp_system") or "mcframe").strip().lower()
    erp_system = erp_raw if erp_raw in ("mcframe", "sap_b1", "excel") else "mcframe"
    if selected_profile:
        profile_erp = str(selected_profile.get("erp_system") or "").strip().lower()
        if profile_erp in ("mcframe", "sap_b1", "excel") and profile_erp != erp_system:
            flash("The selected Customer and System combination does not match.", "error")
            return redirect(url_for("index"))
    if erp_system != "excel" and (not oracle_id or not oracle_pwd or not oracle_schema or not oracle_dsn):
        flash("Please enter ID / PASSWORD / SCHEMA(DB Name) / DNS(Server Name).", "error")
        return redirect(url_for("index"))

    co_cd_stored: str | None = None
    if erp_system == "mcframe":
        co_raw = request.form.get("mcframe_co_cd", "").strip().upper()
        if not co_raw and selected_profile:
            co_raw = str(selected_profile.get("mcframe_co_cd") or "").strip().upper()
        if not co_raw:
            co_raw = (os.environ.get("BRIDGE_MCFRAME_CO_CD") or "J0001").strip().upper()
        if not _MCFRAME_CO_CD_RE.match(co_raw):
            flash(
                "Company CD must be 1 to 20 half-width alphanumeric characters "
                "(when blank, J0001 or BRIDGE_MCFRAME_CO_CD is used).",
                "error",
            )
            return redirect(url_for("index"))
        co_cd_stored = co_raw

    if erp_system == "excel":
        schema_stored = ""
        oracle_dsn = ""
    elif erp_system == "sap_b1":
        if not _SAP_B1_DB_RE.match(oracle_schema):
            flash(
                "Database name (SCHEMA field) must be 1 to 128 characters and "
                "can contain only alphanumeric characters, dots, hyphens, and underscores.",
                "error",
            )
            return redirect(url_for("index"))
        schema_stored = oracle_schema
        try:
            # SAP B1 接続テスト。pyodbc / SQL Server ドライバへの依存は
            # この import 1 箇所に閉じ込めることで、SAP B1 を含めない納品でも
            # アプリ起動に失敗しないようにする。
            from core.erp.sap_b1.connection import connect_sqlserver  # noqa: PLC0415

            conn = connect_sqlserver(
                oracle_dsn, schema_stored, oracle_id, oracle_pwd, timeout=15
            )
            conn.close()
        except Exception as exc:  # noqa: BLE001
            flash(f"Connection failed: {exc}", "error")
            session["oracle_connected"] = False
            return redirect(url_for("index"))
    else:
        if not _SCHEMA_RE.match(oracle_schema):
            flash("Schema must start with a letter and contain only alphanumeric characters and underscores, up to 30 characters.", "error")
            return redirect(url_for("index"))
        schema_stored = oracle_schema.upper()
        try:
            # mcframe (Oracle) 接続テスト。oracledb / Oracle Client への依存は
            # ここでだけ発生するため、Excel/SAP B1 のみの納品でも import エラーで
            # アプリが起動しなくなることはない。
            _ensure_oracledb_initialized()
            import oracledb  # noqa: PLC0415

            conn = oracledb.connect(
                user=oracle_id, password=oracle_pwd, dsn=oracle_dsn
            )
            conn.close()
        except Exception as exc:  # noqa: BLE001
            flash(f"Connection failed: {exc}", "error")
            session["oracle_connected"] = False
            return redirect(url_for("index"))

    # 接続成功時のみ保持（モーダルは前回成功時の値をデフォルト表示）
    session.permanent = True
    session["erp_system"] = erp_system
    session["oracle_user"] = oracle_id if erp_system != "excel" else ""
    session["oracle_password"] = oracle_pwd if erp_system != "excel" else ""
    session["oracle_schema"] = schema_stored
    session["oracle_dsn"] = oracle_dsn
    session["oracle_connected"] = True
    session["customer_id"] = customer_id
    if erp_system == "mcframe" and co_cd_stored:
        session["mcframe_co_cd"] = co_cd_stored
    else:
        session.pop("mcframe_co_cd", None)
    flash("Connection successful.", "success")
    return redirect(url_for("index"))


def _csv_response(csv_data: str, filename: str) -> Response:
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def _require_connection_or_redirect():
    """DISCONNECTED 状態のままなら ``flash`` して index へ redirect する。

    CSV ダウンロード系ハンドラの先頭で呼び、戻り値が ``None`` でなければ
    そのまま return する想定。Excel のように接続不要な ERP は ``ping()`` が
    no-op のため常にここを通り抜ける。Line Cycle Time Master は接続不要なので
    このガードを呼ばない。

    ``fetch()`` は 302 を自動追従するため、常に redirect すると HTML が
    Blob 化されてしまう。``Sec-Fetch-Dest: empty``（典型的な fetch）のときは
    本文付き 401 を返し、クライアントで判別できるようにする。
    """

    if _verify_connection_alive():
        return None
    msg = "ERP is not connected. Please connect from Connect first."
    flash(msg, "error")
    dest = (request.headers.get("Sec-Fetch-Dest") or "").lower()
    mode = (request.headers.get("Sec-Fetch-Mode") or "").lower()
    if dest == "empty" and mode in ("cors", "same-origin", "no-cors"):
        return Response(
            msg + "\n",
            401,
            mimetype="text/plain; charset=utf-8",
        )
    return redirect(url_for("index"))


def _download_error_response(message: str) -> Response:
    """POST /download/* が失敗したとき。fetch 由来なら redirect せず本文を返す。"""
    flash(message, "error")
    dest = (request.headers.get("Sec-Fetch-Dest") or "").lower()
    mode = (request.headers.get("Sec-Fetch-Mode") or "").lower()
    if dest == "empty" and mode in ("cors", "same-origin", "no-cors"):
        return Response(
            message + "\n",
            400,
            mimetype="text/plain; charset=utf-8",
        )
    return redirect(url_for("index"))


def _bridge_kind_download_allowed(kind: str) -> bool:
    """TRANSACTION で ``disabled`` のボタンに対応する kind は POST でも拒否する。"""

    for btn in get_customer(session.get("customer_id")).bridge_buttons():
        if btn.kind == kind:
            return not btn.disabled
    return True


def _sanitize_item_table_rows(rows: list[tuple]) -> list[tuple]:
    """Item table CSV: Asprova 連携都合で ITM_NM 内のカンマを ``#`` に置換する。"""
    out: list[tuple] = []
    for r in rows:
        if not r:
            out.append(r)
            continue
        cells = list(r)
        if len(cells) > 1 and cells[1] is not None:
            cells[1] = str(cells[1]).replace(",", "#")
        out.append(tuple(cells))
    return out


@app.route("/download/integrated", methods=["POST"])
def download_integrated():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    try:
        service = _get_erp_service()
        records = service.fetch_integrated_records(
            upload=request.files.get("source_excel")
        )
        records = _apply_cycle_time_to_integrated_records(records)
        rows = [tuple(d[h] for h in INTEGRATED_HEADERS) for d in records]
        csv_data = rows_to_csv(rows, INTEGRATED_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "integrated_master.csv")


@app.route("/download/item-table", methods=["POST"])
def download_item_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    try:
        service = _get_erp_service()
        rows = _sanitize_item_table_rows(
            service.fetch_item_rows(upload=request.files.get("source_excel"))
        )
        csv_data = rows_to_csv(rows, ITEM_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "item_table.csv")


@app.route("/download/order-table", methods=["POST"])
def download_order_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    if not _bridge_kind_download_allowed("order"):
        return _download_error_response(
            "This download is not available for the current customer."
        )
    try:
        service = _get_erp_service()
        rows = service.fetch_order_rows(upload=request.files.get("source_excel"))
        csv_data = rows_to_csv(rows, ORDER_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "order_table.csv")


@app.route("/download/prd-plan-table", methods=["POST"])
def download_prd_plan_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    try:
        service = _get_erp_service()
        rows = service.fetch_prd_plan_rows(
            upload=request.files.get("source_excel")
        )
        csv_data = rows_to_csv(rows, ORDER_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "prd_plan_table.csv")


@app.route("/download/resource-table", methods=["POST"])
def download_resource_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    try:
        service = _get_erp_service()
        rows = service.fetch_resource_rows(
            upload=request.files.get("source_excel"),
            sort_order_map=_load_sort_order_map_from_sqlite(),
        )
        csv_data = rows_to_csv(rows, RESOURCE_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "resource_table.csv")


@app.route("/download/inventory-table", methods=["POST"])
def download_inventory_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    if not _bridge_kind_download_allowed("inventory"):
        return _download_error_response(
            "This download is not available for the current customer."
        )
    try:
        service = _get_erp_service()
        rows = service.fetch_inventory_rows(
            upload=request.files.get("source_excel")
        )
        csv_data = rows_to_csv(rows, INVENTORY_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "inventory_table.csv")


@app.route("/download/inventory-wip-table", methods=["POST"])
def download_inventory_wip_table():
    guard = _require_connection_or_redirect()
    if guard is not None:
        return guard
    try:
        service = _get_erp_service()
        rows = service.fetch_inventory_wip_rows(
            upload=request.files.get("source_excel")
        )
        csv_data = rows_to_csv(rows, INVENTORY_TABLE_HEADERS)
    except NotSupportedError as exc:
        return _download_error_response(str(exc))
    except Exception as exc:  # noqa: BLE001
        return _download_error_response(f"An error occurred: {exc}")
    return _csv_response(csv_data, "inventory_wip_table.csv")


@app.route("/resource-cycle-times", methods=["GET"])
def resource_cycle_times():
    source_name = str(session.get(RESOURCE_CYCLE_SOURCE_SESSION_KEY) or "").strip()
    active_co_cd = _resource_cycle_scope_key()
    try:
        view_rows = _fetch_resource_cycle_rows()
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to load cycle time SQLite data: {exc}", "error")
        view_rows = []
    return render_template(
        "bridge_cycle_time.html",
        oracle_connected=_verify_connection_alive(),
        erp_system=session.get("erp_system") or "mcframe",
        customer_profiles=_customer_profile_options(),
        selected_customer_id=str(session.get("customer_id") or ""),
        active_co_cd=active_co_cd,
        source_name=source_name,
        rows=view_rows,
    )


@app.route("/resource-cycle-times", methods=["POST"])
def save_resource_cycle_times():
    try:
        rows = _fetch_resource_cycle_rows()
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to load cycle time SQLite data: {exc}", "error")
        return redirect(url_for("index", show_line_cycle="1"))
    try:
        if not rows:
            flash("There is nothing to save. Please import a Resource CSV first.", "error")
            return redirect(url_for("index", show_line_cycle="1"))
        updates: dict[int, dict[str, str | int]] = {}
        for row in rows:
            row_id = int(row["id"])
            cycle_time = (request.form.get(f"cycle_time_{row_id}") or "").strip()
            sort_raw = (request.form.get(f"sort_order_{row_id}") or "").strip()
            if sort_raw == "":
                sort_order = int(row.get("sort_order") or 0)
            else:
                try:
                    sort_order = int(sort_raw)
                except ValueError as exc:
                    raise RuntimeError(f"Sort Order must be an integer (ID={row_id}).") from exc
            updates[row_id] = {"cycle_time": cycle_time, "sort_order": sort_order}
        _update_resource_cycle_rows_by_id(updates)
        flash("Cycle time data has been saved.", "success")
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to save cycle time SQLite data: {exc}", "error")
    return redirect(url_for("index", show_line_cycle="1"))


@app.route("/resource-cycle-times/delete-row", methods=["POST"])
def delete_resource_cycle_time_row():
    row_id_raw = (request.form.get("row_id") or "").strip()
    if not row_id_raw.isdigit():
        flash("The row number to delete is invalid.", "error")
        return redirect(url_for("index", show_line_cycle="1"))
    row_id = int(row_id_raw)
    try:
        deleted = _delete_resource_cycle_row_by_id(row_id)
        if deleted <= 0:
            raise RuntimeError("The row to delete was not found.")
        flash("The row has been deleted.", "success")
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to delete the row: {exc}", "error")
    return redirect(url_for("index", show_line_cycle="1"))


@app.route("/resource-cycle-times/add-row", methods=["POST"])
def add_resource_cycle_time_row():
    line_code = (request.form.get("new_line_code") or "").strip()
    line_name = (request.form.get("new_line_name") or "").strip()
    cycle_time = (request.form.get("new_cycle_time") or "").strip()
    sort_raw = (request.form.get("new_sort_order") or "").strip()
    if not line_code:
        flash("Line Code is required.", "error")
        return redirect(url_for("index", show_line_cycle="1"))
    try:
        sort_order = int(sort_raw) if sort_raw else 0
    except ValueError:
        flash("Sort Order must be an integer.", "error")
        return redirect(url_for("index", show_line_cycle="1"))
    try:
        _insert_resource_cycle_row(
            line_code=line_code,
            line_name=line_name,
            cycle_time=cycle_time,
            sort_order=sort_order,
        )
        flash(f"Row added: {line_code}", "success")
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to add the row: {exc}", "error")
    return redirect(url_for("index", show_line_cycle="1"))


@app.route("/resource-cycle-times/import", methods=["POST"])
def import_resource_cycle_times():
    upload = request.files.get("resource_csv")
    if not upload or not (upload.filename or "").strip():
        flash("Please select the Resource CSV file to import.", "error")
        return redirect(url_for("index", show_line_cycle="1"))
    try:
        raw = upload.read()
        if not raw:
            raise RuntimeError("The uploaded file is empty.")
        rows = _read_resource_cycle_rows_from_bytes(raw)
        file_name = secure_filename(upload.filename or "resource_table.csv")
        if not file_name:
            file_name = "resource_table.csv"
        if not file_name.lower().endswith(".csv"):
            file_name = f"{file_name}.csv"
        _replace_resource_cycle_rows(rows, file_name, _resource_cycle_scope_key())
        session[RESOURCE_CYCLE_SOURCE_SESSION_KEY] = file_name
        flash(f"Resource CSV imported and saved to SQLite: {file_name}", "success")
    except Exception as exc:  # noqa: BLE001
        flash(f"Failed to import Resource CSV: {exc}", "error")
    return redirect(url_for("index", show_line_cycle="1"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5001")), debug=True)

