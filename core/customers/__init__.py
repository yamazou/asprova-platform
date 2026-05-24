"""顧客別ロジックの公開エントリポイント。

使い方:

    >>> from core.customers import get_customer
    >>> strategy = get_customer(session.get("customer_id"))
    >>> if strategy.psi_split_by_customer():
    ...     ...

新顧客を追加する場合は、

    1. ``core/customers/<id>.py`` に ``CustomerStrategy`` 派生クラスを作成
    2. 下記 ``_REGISTRY`` に ``"<id>": ("core.customers.<id>", "<ClassName>")`` を登録
    3. ``config/bridge_customers.py`` の ``BRIDGE_CUSTOMERS`` にも
       接続情報用エントリを追加

の 3 点で完結する。``apps/`` 配下のハンドラ／テンプレートに
顧客 ID 直書きの ``if`` 分岐を増やさないこと。

レジストリは ``("<module path>", "<class name>")`` 形式の **lazy import**。
これにより、PHC 納品で ``core/customers/peb.py`` を物理削除しても
``get_customer("phc")`` は壊れず、PEB を呼び出さない限り import も発生しない。
"""

from __future__ import annotations

import importlib
from typing import Optional

from .base import (
    BridgeButton,
    CustomerStrategy,
    CustomerView,
    DefaultCustomer,
    PsiRowDefinition,
)


# customer_id -> (module path, class name) の lazy import レジストリ。
_REGISTRY: dict[str, tuple[str, str]] = {
    "nci": ("core.customers.nci", "NciCustomer"),
    "phc": ("core.customers.phc", "PhcCustomer"),
    "peb": ("core.customers.peb", "PebCustomer"),
    "sip": ("core.customers.sip", "SipCustomer"),
}


def get_customer(customer_id: Optional[str]) -> CustomerStrategy:
    """顧客 ID から Strategy を取得する。未登録の場合は ``DefaultCustomer``。

    NCI は ``NciCustomer`` で Excel マスタ出力と KOITO/HPM スケジュールを提供する。
    SW / DEMO のような他顧客はレジストリ非登録のまま
    ``DefaultCustomer`` を共有する。
    """

    key = (customer_id or "").strip().lower()
    entry = _REGISTRY.get(key)
    if entry is None:
        return DefaultCustomer()
    module_path, class_name = entry
    module = importlib.import_module(module_path)
    cls = getattr(module, class_name)
    return cls()


__all__ = [
    "BridgeButton",
    "CustomerStrategy",
    "CustomerView",
    "DefaultCustomer",
    "PsiRowDefinition",
    "get_customer",
]
