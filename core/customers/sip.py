"""SIP customer behavior."""

from __future__ import annotations

from .base import CustomerStrategy


class SipCustomer(CustomerStrategy):
    id = "sip"
    label = "SIP"
