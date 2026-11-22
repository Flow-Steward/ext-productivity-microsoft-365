"""Shared helpers of the Teams package: closed input validation, Graph paths and HTML text."""

from __future__ import annotations

import html
import re
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from html.parser import HTMLParser
from typing import Any
from urllib.parse import quote, urlsplit

from m365.errors import M365Error

#: Team ids are the Microsoft 365 group's object id.
_GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
#: Channel and chat ids: ``19:...@thread.tacv2``, ``19:..._...@unq.gbl.spaces``, ``19:meeting_...``.
_THREAD_ID = re.compile(r"^19:[A-Za-z0-9._@=-]{1,500}$")
#: Message ids are numeric timestamps today; accept a URL-safe superset.
_MESSAGE_ID = re.compile(r"^[A-Za-z0-9_-]{1,128}$")
_EMAIL = re.compile(
    r"^[^@\s<>()\[\],;:\"\\]{1,64}@[A-Za-z0-9](?:[A-Za-z0-9.-]{0,251}[A-Za-z0-9])?\.[A-Za-z]{2,63}$"
)
MAX_MENTIONS = 20
MAX_FILE_LINKS = 10
MAX_BODY_CHARS = 200_000
MAX_SUBJECT_CHARS = 255
MAX_URL_CHARS = 2048


# -- input validation --------------------------------------------------------------------


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
    if required and not result.strip():
        raise M365Error("invalid_payload", f"{field} is required")
    if len(result) > max_length or not _clean(result, allow_newlines=allow_newlines):
        raise M365Error("invalid_payload", f"{field} is invalid")
    return result


def int_input(value: Any, field: str, *, default: int, minimum: int, maximum: int) -> int:
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise M365Error(
            "invalid_payload", f"{field} must be an integer from {minimum} to {maximum}"
        )
    return value


def team_id_input(value: Any, field: str = "team_id") -> str:
    if not isinstance(value, str) or not _GUID.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} must be a Teams team ID")
    return value.strip().lower()


def thread_id_input(value: Any, field: str) -> str:
    """A channel or chat id (``19:...``)."""
    if not isinstance(value, str) or not _THREAD_ID.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} must be a Teams {field.split('_')[0]} ID")
    return value.strip()


def message_id_input(value: Any, field: str = "message_id") -> str:
    if not isinstance(value, str) or not _MESSAGE_ID.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} must be a Teams message ID")
    return value.strip()


def email_input(value: Any, field: str) -> str:
    if not isinstance(value, str) or not _EMAIL.match(value.strip()):
        raise M365Error("invalid_payload", f"{field} must be an email address")
    return value.strip()


def mentions_input(value: Any) -> list[str]:
    """Addresses of users to @mention, de-duplicated case-insensitively."""
    if value is None or value == []:
        return []
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list) or len(value) > MAX_MENTIONS:
        raise M365Error(
            "invalid_payload", f"mentions must be a list of at most {MAX_MENTIONS} addresses"
        )
    result: list[str] = []
    seen: set[str] = set()
    for raw in value:
        address = email_input(raw, "mentions")
        if address.lower() not in seen:
            seen.add(address.lower())
            result.append(address)
    return result


def https_url_input(value: Any, field: str) -> str:
    if not isinstance(value, str) or len(value.strip()) > MAX_URL_CHARS:
        raise M365Error("invalid_payload", f"{field} must be an https URL")
    url = value.strip()
    try:
        parsed = urlsplit(url)
    except ValueError as exc:
        raise M365Error("invalid_payload", f"{field} must be an https URL") from exc
    if (
        parsed.scheme != "https"
        or not parsed.hostname
        or parsed.username
        or not _clean(url, allow_newlines=False)
        or any(char.isspace() for char in url)
    ):
        raise M365Error("invalid_payload", f"{field} must be an https URL")
    return url


def file_links_input(value: Any) -> list[tuple[str, str]]:
    """``[(url, name)]`` from URLs or ``{"url": ..., "name": ...}`` objects."""
    if value is None or value == []:
        return []
    if isinstance(value, str | Mapping):
        value = [value]
    if not isinstance(value, list) or len(value) > MAX_FILE_LINKS:
        raise M365Error(
            "invalid_payload", f"file_links must be a list of at most {MAX_FILE_LINKS} links"
        )
    result: list[tuple[str, str]] = []
    for raw in value:
        if isinstance(raw, Mapping):
            if set(raw) - {"url", "name"}:
                raise M365Error("invalid_payload", "file_links entries allow only url and name")
            url = https_url_input(raw.get("url"), "file_links.url")
            name = text_input(raw.get("name"), "file_links.name", max_length=256)
        else:
            url = https_url_input(raw, "file_links")
            name = ""
        result.append((url, name or url))
    return result


# -- Graph paths and ids -------------------------------------------------------------------


def segment(value: str) -> str:
    return quote(value, safe="@:._=-")


def channel_path(team_id: str, channel_id: str) -> str:
    return f"/teams/{segment(team_id)}/channels/{segment(channel_id)}"


def chat_path(chat_id: str) -> str:
    return f"/chats/{segment(chat_id)}"


def provider_text(value: Any, *, limit: int = 1024) -> str:
    return value[:limit] if isinstance(value, str) else ""


def provider_message_id(value: Any) -> str:
    if not isinstance(value, str) or not _MESSAGE_ID.match(value):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid message ID")
    return value


def provider_thread_id(value: Any) -> str:
    if not isinstance(value, str) or not _THREAD_ID.match(value):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid chat ID")
    return value


@contextmanager
def not_found_as(code: str, message: str) -> Iterator[None]:
    """Re-map the generic ``item_not_found`` to a specific ``*_not_found`` code."""
    try:
        yield
    except M365Error as exc:
        if exc.code == "item_not_found":
            raise M365Error(code, message, http_status=exc.http_status) from exc
        raise


def list_all(
    transport: Any, path: str, *, query: Mapping[str, Any] | None, max_items: int, purpose: str
) -> list[dict[str, Any]]:
    """Follow ``@odata.nextLink`` up to ``max_items`` rows."""
    page = transport.get_json(path, query=query, purpose=purpose)
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
        page = transport.get_json(link, purpose=purpose)


# -- HTML ------------------------------------------------------------------------------------


def text_to_html(value: str) -> str:
    """Plain text as Teams HTML: escaped, with line breaks kept."""
    escaped = html.escape(value, quote=False)
    return escaped.replace("\r\n", "\n").replace("\n", "<br>")


_BLOCK_TAGS = frozenset({"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"})


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style"}:
            self._skip += 1
        elif tag in _BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style"}:
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip:
            self.parts.append(data)


def html_to_text(value: str) -> str:
    parser = _TextExtractor()
    try:
        parser.feed(value)
        parser.close()
    except Exception:
        return re.sub(r"<[^>]*>", "", value).strip()
    raw = "".join(parser.parts).replace("\xa0", " ")
    lines = [line.rstrip() for line in raw.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


__all__ = [
    "MAX_BODY_CHARS",
    "MAX_SUBJECT_CHARS",
    "channel_path",
    "chat_path",
    "email_input",
    "file_links_input",
    "html_to_text",
    "https_url_input",
    "int_input",
    "list_all",
    "mentions_input",
    "message_id_input",
    "not_found_as",
    "provider_message_id",
    "provider_text",
    "provider_thread_id",
    "segment",
    "team_id_input",
    "text_input",
    "text_to_html",
    "thread_id_input",
]
