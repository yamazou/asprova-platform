"""PHC 顧客向けの固有ロジック。

主な差分:
    - PSI ビューを Customer (Export / Internal) 軸で行分割する
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from typing import Any, Callable, Optional

from .base import CustomerStrategy, PsiRowDefinition


class PhcCustomer(CustomerStrategy):
    id = "phc"
    label = "PHC"

    # ------------------------------------------------------------------
    # PSI view
    # ------------------------------------------------------------------

    def psi_split_by_customer(self) -> bool:
        return True

    def psi_row_definitions(self) -> list[PsiRowDefinition]:
        return [
            PsiRowDefinition("Supply", "P", "Export", 2, "Export"),
            PsiRowDefinition("Supply", "P", "Internal", 0, "Internal"),
            PsiRowDefinition("Demand", "S", "Export", 2, "Export"),
            PsiRowDefinition("Demand", "S", "Internal", 0, "Internal"),
            PsiRowDefinition("Stock", "I", "", 1, None),
        ]

    def build_psi_split_aggs(
        self,
        start_date: date,
        end_date: date,
        db_factory: Callable[[], Any],
    ) -> tuple[dict, dict]:
        """psi_input_instructions / psi_output_instructions から
        Customer (Export/Internal) 単位の集計 dict を構築する。

        旧 ``apps/viewer/app.py:_build_psi_customer_split_aggs_for_month``
        をここへ移動。app 層から DB ハンドルを受け取ることで
        ``apps/viewer`` への逆依存を作らない。
        """

        start_s = start_date.strftime("%Y-%m-%d")
        end_s = end_date.strftime("%Y-%m-%d")

        conn = db_factory()
        try:
            out_rows = conn.execute(
                """
                SELECT item_code, inst_time, quantity, operation_code, customer
                FROM psi_output_instructions
                WHERE inst_time >= ? AND inst_time < ?
                """,
                (start_s, end_s),
            ).fetchall()
            in_rows = conn.execute(
                """
                SELECT item_code, inst_time, quantity, operation_code, customer
                FROM psi_input_instructions
                WHERE inst_time >= ? AND inst_time < ?
                """,
                (start_s, end_s),
            ).fetchall()
            sched_rows = conn.execute(
                """
                SELECT
                    operation_code,
                    COALESCE(NULLIF(TRIM(machine_name), ''), NULLIF(TRIM(actual_resource), ''), 'Unknown') AS machine_name
                FROM schedules
                WHERE start_time >= ? AND start_time < ?
                  AND operation_code IS NOT NULL
                  AND TRIM(operation_code) <> ''
                """,
                (start_s, end_s),
            ).fetchall()
        finally:
            conn.close()

        op_to_machine_counts = defaultdict(lambda: defaultdict(int))
        for sr in sched_rows:
            op_code = _op_link_key(sr["operation_code"])
            machine_name = str(sr["machine_name"] or "").strip() or "Unknown"
            if not op_code:
                continue
            op_to_machine_counts[op_code][machine_name] += 1

        def _machine_dim(op_raw) -> str:
            op_raw_txt = str(op_raw or "").strip()
            op_code = _op_link_key(op_raw_txt)
            if not op_code:
                return "Unknown"
            counts = op_to_machine_counts.get(op_code)
            if counts:
                return max(counts.items(), key=lambda kv: kv[1])[0]
            return op_raw_txt or op_code

        supply_split: dict = defaultdict(float)
        for r in out_rows:
            item_key = str(r["item_code"] or "").strip()
            if not item_key:
                continue
            day = _day_from_inst_time(r["inst_time"])
            if day is None:
                continue
            qty = float(r["quantity"] or 0)
            # CUSTOMER アイテムは出荷扱いのため符号を反転
            if item_key.upper() == "CUSTOMER" and qty > 0:
                qty = -abs(qty)
            machine = _machine_dim(r["operation_code"])
            bucket = _customer_bucket(r["customer"])
            if bucket is None:
                continue
            supply_split[(item_key, day, bucket, machine)] += qty

        demand_split: dict = defaultdict(float)
        for r in in_rows:
            item_key = str(r["item_code"] or "").strip()
            if not item_key:
                continue
            day = _day_from_inst_time(r["inst_time"])
            if day is None:
                continue
            qty = float(r["quantity"] or 0)
            if qty <= 0:
                continue
            machine = _machine_dim(r["operation_code"])
            bucket = _customer_bucket(r["customer"])
            if bucket is None:
                continue
            demand_split[(item_key, day, bucket, machine)] += qty

        return supply_split, demand_split


# ---------------------------------------------------------------------------
# module-private helpers
# ---------------------------------------------------------------------------


def _day_from_inst_time(inst_time):
    if not inst_time:
        return None
    s = str(inst_time).strip()
    if len(s) < 8:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def _op_link_key(raw_op) -> str:
    op = str(raw_op or "").strip()
    if not op:
        return ""
    return op.split(":", 1)[0].strip() if ":" in op else op


def _customer_bucket(raw) -> Optional[str]:
    v = str(raw or "").strip().lower()
    if not v:
        return None
    if "export" in v:
        return "Export"
    if "internal" in v:
        return "Internal"
    return None
