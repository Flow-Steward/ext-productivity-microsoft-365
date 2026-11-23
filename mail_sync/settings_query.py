"""The new-email trigger page's view of one mailbox (v0.2.7).

The page picks a mail connection at the top (ADR 0067). Its form loads this query
and saves through ``save_trigger_settings`` (Core's
``project_connection_ingress_patch``). The values come from the same place the
poll reads them, so the form always shows what the poll will do.
"""

from __future__ import annotations

from typing import Any

from m365.errors import M365Error, error_from_exception

from .poll import read_poll_inputs
from .state import LAST_SYNC_KEY, MODE_OFF, MODE_ON

QUERY_ID = "mail_trigger_settings"


def handle_trigger_settings_query(payload: dict[str, Any]) -> dict[str, Any]:
    try:
        inputs = read_poll_inputs(payload)
    except M365Error as exc:
        return error_from_exception(exc, effectful=False)
    mode = MODE_ON if inputs.mode == MODE_ON else MODE_OFF
    last_check = str(inputs.ingress.get(LAST_SYNC_KEY) or "")
    return {
        "ok": True,
        "values": {"mail_trigger_mode": mode, "mail_trigger_folder": inputs.folder},
        "items": [
            {
                "trigger": "On" if mode == MODE_ON else "Off",
                "folder": inputs.folder,
                "last_check": last_check,
            }
        ],
    }


__all__ = ["QUERY_ID", "handle_trigger_settings_query"]
