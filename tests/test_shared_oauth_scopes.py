"""Delegated connection types must not require scopes Microsoft never lists as granted."""

from __future__ import annotations

from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[1]
#: Requested, but absent from the token response's `scope` (Microsoft identity platform).
NEVER_LISTED = {"offline_access"}


def _oauth_types() -> list[dict]:
    data = yaml.safe_load((ROOT / "contracts" / "connection_types.yaml").read_text())
    return [t for t in data["connection_types"] if (t.get("auth") or {}).get("type") == "oauth2"]


def test_offline_access_is_requested_but_not_required() -> None:
    types = _oauth_types()
    assert types
    for connection_type in types:
        scopes = connection_type["auth"]["scopes"]
        assert not NEVER_LISTED & set(scopes["required"]), connection_type["connection_type_id"]
        assert "offline_access" in scopes["optional"], connection_type["connection_type_id"]
