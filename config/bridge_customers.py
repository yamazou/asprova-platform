from __future__ import annotations

import os


def _env(name: str, default: str) -> str:
    return (os.environ.get(name) or default).strip()


# 顧客追加はこの辞書へ 1 エントリ追加するだけで対応可能。
# 値は Connect モーダルの初期値として使われ、必要に応じてユーザーが上書きできます。
BRIDGE_CUSTOMERS: dict[str, dict[str, str]] = {
    "demo": {
        "label": "DEMO",
        "erp_system": "mcframe",
        "oracle_id": _env("BRIDGE_DEMO_USER", "demo"),
        "oracle_pwd": _env("BRIDGE_DEMO_PASSWORD", "demo"),
        "oracle_schema": _env("BRIDGE_DEMO_SCHEMA", "demo"),
        "oracle_dsn": _env("BRIDGE_DEMO_DSN", "orcl"),
        "mcframe_co_cd": _env("BRIDGE_DEMO_CO_CD", "J0001"),
    },
    "nci": {
        "label": "NCI",
        "erp_system": "mcframe",
        "oracle_id": _env("BRIDGE_NCI_USER", "nci"),
        "oracle_pwd": _env("BRIDGE_NCI_PASSWORD", "nci"),
        "oracle_schema": _env("BRIDGE_NCI_SCHEMA", "nci"),
        "oracle_dsn": _env("BRIDGE_NCI_DSN", "orcl"),
        "mcframe_co_cd": _env("BRIDGE_NCI_CO_CD", "NCI"),
    },
    "sw": {
        "label": "SW",
        "erp_system": "sap_b1",
        "oracle_id": _env("BRIDGE_SW_USER", "sa"),
        "oracle_pwd": _env("BRIDGE_SW_PASSWORD", "riniradi66"),
        "oracle_schema": _env("BRIDGE_SW_DATABASE", "SW"),
        "oracle_dsn": _env("BRIDGE_SW_SERVER", "LAPTOP-4ST122V3\SQLEXPRESS"),
        "mcframe_co_cd": "",
    },
    "peb": {
        "label": "PEB",
        "erp_system": "excel",
        "oracle_id": "",
        "oracle_pwd": "",
        "oracle_schema": "",
        "oracle_dsn": "",
        "mcframe_co_cd": "",
        "excel_base_dir": _env("BRIDGE_PEB_EXCEL_DIR", "data/bridge_excel/peb"),
        "excel_integrated_file": _env("BRIDGE_PEB_INTEGRATED_FILE", "integrated_master.xlsx"),
        "excel_item_file": _env("BRIDGE_PEB_ITEM_FILE", "item_table.xlsx"),
        "excel_order_file": _env("BRIDGE_PEB_ORDER_FILE", "order_table.xlsx"),
        "excel_prd_plan_file": _env("BRIDGE_PEB_PRD_PLAN_FILE", "prd_plan_table.xlsx"),
        "excel_resource_file": _env("BRIDGE_PEB_RESOURCE_FILE", "resource_table.xlsx"),
        "excel_inventory_file": _env("BRIDGE_PEB_INVENTORY_FILE", "inventory_table.xlsx"),
        "excel_inventory_wip_file": _env("BRIDGE_PEB_INVENTORY_WIP_FILE", "inventory_wip_table.xlsx"),
    },
    "phc": {
        "label": "PHC",
        "erp_system": "excel",
        "oracle_id": "",
        "oracle_pwd": "",
        "oracle_schema": "",
        "oracle_dsn": "",
        "mcframe_co_cd": "",
        "excel_base_dir": _env("BRIDGE_PHC_EXCEL_DIR", "data/bridge_excel/phc"),
        "excel_integrated_file": _env("BRIDGE_PHC_INTEGRATED_FILE", "integrated_master.xlsx"),
        "excel_item_file": _env("BRIDGE_PHC_ITEM_FILE", "item_table.xlsx"),
        "excel_order_file": _env("BRIDGE_PHC_ORDER_FILE", "order_table.xlsx"),
        "excel_prd_plan_file": _env("BRIDGE_PHC_PRD_PLAN_FILE", "prd_plan_table.xlsx"),
        "excel_resource_file": _env("BRIDGE_PHC_RESOURCE_FILE", "resource_table.xlsx"),
        "excel_inventory_file": _env("BRIDGE_PHC_INVENTORY_FILE", "inventory_table.xlsx"),
        "excel_inventory_wip_file": _env("BRIDGE_PHC_INVENTORY_WIP_FILE", "inventory_wip_table.xlsx"),
    },
    "sip": {
        "label": "SIP",
        "erp_system": "excel",
        "oracle_id": "",
        "oracle_pwd": "",
        "oracle_schema": "",
        "oracle_dsn": "",
        "mcframe_co_cd": "",
        "excel_base_dir": _env("BRIDGE_SIP_EXCEL_DIR", "data/bridge_excel/sip"),
        "excel_integrated_file": _env("BRIDGE_SIP_INTEGRATED_FILE", "integrated_master.xlsx"),
        "excel_item_file": _env("BRIDGE_SIP_ITEM_FILE", "item_table.xlsx"),
        "excel_order_file": _env("BRIDGE_SIP_ORDER_FILE", "order_table.xlsx"),
        "excel_prd_plan_file": _env("BRIDGE_SIP_PRD_PLAN_FILE", "prd_plan_table.xlsx"),
        "excel_resource_file": _env("BRIDGE_SIP_RESOURCE_FILE", "resource_table.xlsx"),
        "excel_inventory_file": _env("BRIDGE_SIP_INVENTORY_FILE", "inventory_table.xlsx"),
        "excel_inventory_wip_file": _env("BRIDGE_SIP_INVENTORY_WIP_FILE", "inventory_wip_table.xlsx"),
    },
}

