"""Shared helpers of the mail package: closed input validation and the mailbox client."""

from __future__ import annotations

import re
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from datetime import UTC, datetime
from typing import Any
from urllib.parse import quote

from m365.connection import ConnectionContext, mailbox_base
from m365.errors import M365Error
from m365.transport import DEFAULT_RESPONSE_LIMIT, IMMUTABLE_ID_PREFER, GraphResponse

#: Outlook well-known folder names accepted wherever a folder is expected.
WELL_KNOWN_FOLDERS: dict[str, str] = {
    "inbox": "Inbox",
    "drafts": "Drafts",
    "sentitems": "Sent Items",
    "deleteditems": "Deleted Items",
    "archive": "Archive",
    "junkemail": "Junk Email",
    "outbox": "Outbox",
    "scheduled": "Scheduled",
    "conversationhistory": "Conversation History",
    "msgfolderroot": "Top of Information Store",
}
#: Graph ids (immutable or regular) are URL-safe base64 with optional padding.
_GRAPH_ID = re.compile(r"^[A-Za-z0-9_=-]{1,1024}$")
_EMAIL = re.compile(
    r"^[^@\s<>()\[\],;:\"\\]{1,64}@[A-Za-z0-9](?:[A-Za-z0-9.-]{0,251}[A-Za-z0-9])?\.[A-Za-z]{2,63}$"
)
MAX_MESSAGE_IDS = 100
MAX_RECIPIENTS_TOTAL = 500
BODY_PREFER_TEXT = 'outlook.body-content-type="text"'
BODY_PREFER_HTML = 'outlook.body-content-type="html"'


# -- input validation --------------------------------------------------------


def _clean(value: str, *, allow_newlines: bool) -> bool:
    for char in value:
        code = ord(char)
        if code == 127 or (code < 32 and not (allow_newlines and char in "\r\n\t")):
            return False
    return True


def text_input(
    value: Any,
    field: str,
    *,
    max_length: int,
    required: bool = False,
    allow_newlines: bool = False,
    strip: bool = True,
) -> str:
    if value is None or value == "":
        if required:
            raise M365Error("invalid_payload", f"{field} is required")
        return ""
    if not isinstance(value, str):
        raise M365Error("invalid_payload", f"{field} must be text")
    result = value.strip() if strip else value
    if required and not result:
        raise M365Error("invalid_payload", f"{field} is required")
    if len(result) > max_length or not _clean(result, allow_newlines=allow_newlines):
        raise M365Error("invalid_payload", f"{field} is invalid")
    return result


def bool_input(value: Any, field: str, *, default: bool | None = None) -> bool | None:
    if value is None:
        return default
    if not isinstance(value, bool):
        raise M365Error("invalid_payload", f"{field} must be true or false")
    return value


def int_input(value: Any, field: str, *, default: int, minimum: int, maximum: int) -> int:
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise M365Error(
            "invalid_payload", f"{field} must be an integer from {minimum} to {maximum}"
        )
    return value


def graph_id(value: Any, field: str) -> str:
    if not isinstance(value, str) or not _GRAPH_ID.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} is not a valid Microsoft Graph ID")
    return value.strip()


def provider_id(value: Any) -> str:
    """An id Graph returned; anything unexpected is a provider fault, not caller input."""
    if not isinstance(value, str) or not _GRAPH_ID.match(value):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid ID")
    return value


def folder_ref(value: Any, field: str, *, default: str = "") -> str:
    """A well-known folder name (lower case) or a folder id."""
    if value is None or value == "":
        if default:
            return default
        raise M365Error("invalid_payload", f"{field} is required")
    if not isinstance(value, str):
        raise M365Error("invalid_payload", f"{field} must be text")
    candidate = value.strip()
    if candidate.lower() in WELL_KNOWN_FOLDERS:
        return candidate.lower()
    return graph_id(candidate, field)


def message_ids_input(value: Any) -> list[str]:
    if not isinstance(value, list):
        raise M365Error("invalid_payload", "message_ids must be a list")
    result: list[str] = []
    for raw in value:
        item = graph_id(raw, "message_ids")
        if item not in result:
            result.append(item)
    if not 1 <= len(result) <= MAX_MESSAGE_IDS:
        raise M365Error("invalid_payload", f"message_ids must contain 1 to {MAX_MESSAGE_IDS} IDs")
    return result


def email_address(value: Any, field: str) -> str:
    if not isinstance(value, str) or not _EMAIL.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} contains an invalid email address")
    return value.strip()


def recipients_input(
    value: Any, field: str, *, maximum: int = MAX_RECIPIENTS_TOTAL
) -> list[dict[str, Any]]:
    """Recipients as Graph ``recipient`` objects; accepts addresses or ``{address, name}``."""
    if value is None:
        return []
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list) or len(value) > maximum:
        raise M365Error("invalid_payload", f"{field} must be a list of at most {maximum} addresses")
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for raw in value:
        if isinstance(raw, Mapping):
            if set(raw) - {"address", "name"}:
                raise M365Error("invalid_payload", f"{field} entries allow only address and name")
            address = email_address(raw.get("address"), field)
            name = text_input(raw.get("name"), f"{field}.name", max_length=256)
        else:
            address = email_address(raw, field)
            name = ""
        if address.lower() in seen:
            continue
        seen.add(address.lower())
        email: dict[str, str] = {"address": address}
        if name:
            email["name"] = name
        result.append({"emailAddress": email})
    return result


def datetime_input(value: Any, field: str) -> str:
    """An ISO 8601 date-time with a time zone, normalised to UTC ``...Z``."""
    if value is None or value == "":
        return ""
    if not isinstance(value, str) or len(value) > 64:
        raise M365Error("invalid_payload", f"{field} must be an ISO 8601 date-time")
    try:
        parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError as exc:
        raise M365Error("invalid_payload", f"{field} must be an ISO 8601 date-time") from exc
    if parsed.tzinfo is None:
        raise M365Error("invalid_payload", f"{field} must include a time zone")
    return parsed.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def odata_string(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


# -- provider output ----------------------------------------------------------


def address_out(raw: Any) -> dict[str, str] | None:
    if not isinstance(raw, Mapping):
        return None
    email = raw.get("emailAddress")
    if not isinstance(email, Mapping):
        return None
    return {"address": str(email.get("address") or ""), "name": str(email.get("name") or "")}


def addresses_out(raw: Any) -> list[dict[str, str]]:
    if not isinstance(raw, list):
        return []
    return [row for row in (address_out(item) for item in raw[:MAX_RECIPIENTS_TOTAL]) if row]


def flag_status_out(raw: Any) -> str:
    if isinstance(raw, Mapping):
        status = str(raw.get("flagStatus") or "")
        if status in {"flagged", "complete", "notFlagged"}:
            return status
    return "notFlagged"


@contextmanager
def not_found_as(code: str, message: str) -> Iterator[None]:
    """Re-map the generic ``item_not_found`` to a specific ``*_not_found`` code."""
    try:
        yield
    except M365Error as exc:
        if exc.code == "item_not_found":
            raise M365Error(code, message, http_status=exc.http_status) from exc
        raise


def without_ambiguity(
    exc: M365Error, *, code: str | None = None, message: str | None = None
) -> M365Error:
    """A copy that states nothing external happened, with the provider retry signal."""
    return M365Error(
        code or exc.code,
        message or exc.message,
        http_status=exc.http_status,
        retry_after_seconds=exc.retry_after_seconds,
        ambiguous=False,
        retryable=True if exc.ambiguous else exc.retryable,
        failure_class="provider" if exc.ambiguous else exc.failure_class,
    )


# -- mailbox client -------------------------------------------------------------


def segment(value: str) -> str:
    return quote(value, safe="=-_")


class Mailbox:
    """Graph calls scoped to the connection's mailbox, always with immutable ids."""

    def __init__(self, transport: Any, base: str) -> None:
        self.transport = transport
        self.base = base

    @classmethod
    def for_connection(cls, transport: Any, connection: ConnectionContext) -> Mailbox:
        return cls(transport, mailbox_base(connection))

    @staticmethod
    def headers(*prefer: str, extra: Mapping[str, str] | None = None) -> dict[str, str]:
        headers = {"Prefer": ", ".join((IMMUTABLE_ID_PREFER, *prefer))}
        headers.update(extra or {})
        return headers

    def get(
        self,
        path: str,
        *,
        query: Mapping[str, Any] | None = None,
        prefer: tuple[str, ...] = (),
        response_limit: int = DEFAULT_RESPONSE_LIMIT,
        purpose: str,
    ) -> dict[str, Any]:
        return self.transport.request(
            "GET",
            f"{self.base}{path}",
            query=query,
            headers=self.headers(*prefer),
            response_limit=response_limit,
            retry="read",
            purpose=purpose,
        ).json()

    def get_bytes(self, path: str, *, response_limit: int, purpose: str) -> bytes:
        return self.transport.request(
            "GET",
            f"{self.base}{path}",
            headers=self.headers(extra={"Accept": "*/*"}),
            response_limit=response_limit,
            retry="read",
            purpose=purpose,
        ).body

    def next_page(self, link: str, *, prefer: tuple[str, ...] = (), purpose: str) -> dict[str, Any]:
        return self.transport.request(
            "GET", link, headers=self.headers(*prefer), retry="read", purpose=purpose
        ).json()

    def write(
        self,
        method: str,
        path: str,
        *,
        json_body: Mapping[str, Any] | None = None,
        retry: str,
        purpose: str,
        ok_statuses: tuple[int, ...] = (200, 201, 202, 204),
    ) -> GraphResponse:
        if method.upper() == "GET":
            raise ValueError("write() is for writes")
        return self.transport.request(
            method,
            f"{self.base}{path}",
            json_body=json_body,
            # Graph answers a bodiless POST (e.g. ``/send``) with 411 unless Content-Length is 0.
            raw_body=None if json_body is not None else b"",
            headers=self.headers(),
            retry=retry,
            purpose=purpose,
            ok_statuses=ok_statuses,
        )

    def list_all(
        self,
        path: str,
        *,
        query: Mapping[str, Any] | None = None,
        max_items: int,
        purpose: str,
    ) -> list[dict[str, Any]]:
        """Follow ``@odata.nextLink`` up to ``max_items`` rows."""
        page = self.get(path, query=query, purpose=purpose)
        rows: list[dict[str, Any]] = []
        while True:
            values = page.get("value")
            if not isinstance(values, list):
                raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid list")
            rows.extend(dict(row) for row in values if isinstance(row, Mapping))
            if len(rows) > max_items:
                raise M365Error("response_too_large", "Microsoft Graph returned too many items")
            link = page.get("@odata.nextLink")
            if not isinstance(link, str) or not link:
                return rows
            page = self.next_page(link, purpose=purpose)


__all__ = [
    "BODY_PREFER_HTML",
    "BODY_PREFER_TEXT",
    "WELL_KNOWN_FOLDERS",
    "Mailbox",
    "address_out",
    "addresses_out",
    "bool_input",
    "datetime_input",
    "email_address",
    "flag_status_out",
    "folder_ref",
    "graph_id",
    "int_input",
    "message_ids_input",
    "not_found_as",
    "odata_string",
    "provider_id",
    "recipients_input",
    "segment",
    "text_input",
    "without_ambiguity",
]
