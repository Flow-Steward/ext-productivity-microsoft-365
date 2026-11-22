"""Graph calls of the mail sync: delta, the received-date query and link following.

Every call sends ``Prefer: IdType="ImmutableId"`` so message ids survive moves.
Graph carries ``changeType=created`` and ``$select`` of the first delta request
into every following ``nextLink`` and ``deltaLink``; links are followed verbatim.
The links stay in the connection state only and are never returned in results.
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any
from urllib.parse import quote

from m365.errors import M365Error
from m365.transport import IMMUTABLE_ID_PREFER, GraphTransport

from .machine import FetchPlan, Page
from .state import iso

#: Graph page size for delta and the received-date query.
PAGE_SIZE = 100
#: Fields the events need. Delta keeps the ``$select`` of its first request, so
#: the baseline already asks for everything live pages must carry.
SELECT_FIELDS = "id,internetMessageId,from,subject,receivedDateTime,hasAttachments"
PURPOSE = "Outlook new-mail sync"
_LINK_LIMIT = 8192


class MailSyncGraph:
    """The Graph side of :class:`mail_sync.machine.SyncClient`."""

    def __init__(self, transport: GraphTransport, *, mailbox_base: str, folder: str) -> None:
        self._transport = transport
        self._base = mailbox_base
        self._folder_path = f"{mailbox_base}/mailFolders/{quote(folder, safe='')}"

    @staticmethod
    def headers() -> dict[str, str]:
        return {"Prefer": f"{IMMUTABLE_ID_PREFER}, odata.maxpagesize={PAGE_SIZE}"}

    def fetch(self, plan: FetchPlan) -> Page:
        if plan.kind == "delta_start":
            path, query = (
                f"{self._folder_path}/messages/delta",
                {"changeType": "created", "$select": SELECT_FIELDS},
            )
        elif plan.kind == "query_start":
            if plan.query_from is None:
                raise ValueError("query_start needs query_from")
            path, query = (
                f"{self._folder_path}/messages",
                {
                    "$filter": f"receivedDateTime ge {iso(plan.query_from)}",
                    # Ascending, so mail that arrives while paging lands after the
                    # pages already read and skip-based paging stays stable.
                    "$orderby": "receivedDateTime asc",
                    "$select": SELECT_FIELDS,
                    "$top": str(PAGE_SIZE),
                },
            )
        elif plan.kind == "link":
            path, query = plan.link, {}
        else:
            raise ValueError(f"unsupported fetch plan {plan.kind!r}")
        try:
            body = self._transport.get_json(
                path, query=query or None, headers=self.headers(), purpose=PURPOSE
            )
        except M365Error as exc:
            if exc.code == "item_not_found":
                raise M365Error(
                    "folder_not_found",
                    "The mail folder of the trigger was not found",
                    http_status=exc.http_status,
                ) from exc
            raise
        return parse_page(body)

    def mailbox_address(self) -> str:
        """The signed-in user's address, for a delegated connection without ``mailbox``."""
        body = self._transport.get_json(
            self._base,
            query={"$select": "mail,userPrincipalName"},
            headers={"Prefer": IMMUTABLE_ID_PREFER},
            purpose=PURPOSE,
        )
        for key in ("mail", "userPrincipalName"):
            value = body.get(key)
            if isinstance(value, str) and value.strip() and len(value) <= 320:
                return value.strip().lower()
        return ""


def parse_page(body: Mapping[str, Any]) -> Page:
    values = body.get("value")
    if not isinstance(values, list):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid message page")
    return Page(
        messages=tuple(item for item in values if isinstance(item, Mapping)),
        next_link=_link(body.get("@odata.nextLink")),
        delta_link=_link(body.get("@odata.deltaLink")),
    )


def _link(value: Any) -> str:
    if value is None or value == "":
        return ""
    if not isinstance(value, str) or len(value) > _LINK_LIMIT or not value.startswith("https://"):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid page link")
    return value


__all__ = ["PAGE_SIZE", "SELECT_FIELDS", "MailSyncGraph", "parse_page"]
