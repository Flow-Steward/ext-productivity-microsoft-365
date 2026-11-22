"""Closed input validation and Graph path building for drive items.

Item addresses are either an item id (``/items/{id}``) or a drive path
(``/root:/{path}:``). Path segments are percent-encoded one by one; a name that
SharePoint rejects is refused here before any request is made.
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote

from m365.errors import M365Error

#: Characters SharePoint Online and OneDrive for Business refuse in a name.
_INVALID_NAME_CHARS = frozenset('"*:<>?/\\|')
_ITEM_ID = re.compile(r"^[A-Za-z0-9!_.\-]{1,256}$")
_DRIVE_ID = re.compile(r"^[A-Za-z0-9!_\-]{1,256}$")
_EXTENSION = re.compile(r"^[a-z0-9]{1,16}$")
MAX_PATH_CHARS = 400
#: The shared transport maps a 404 with Graph code ResourceNotFound to
#: mailbox_not_found; on drive calls both mean the addressed object is missing.
NOT_FOUND_CODES = frozenset({"item_not_found", "mailbox_not_found"})
MAX_NAME_BYTES = 255


def item_missing(exc: M365Error) -> M365Error:
    """Report any not-found answer about an addressed item as ``item_not_found``."""
    if exc.code in NOT_FOUND_CODES and exc.code != "item_not_found":
        return M365Error(
            "item_not_found", "The file or folder was not found", http_status=exc.http_status
        )
    return exc


def name(value: Any, field: str) -> str:
    if not isinstance(value, str):
        raise M365Error("invalid_payload", f"{field} must be text")
    if (
        not value
        or value != value.strip()
        or len(value.encode("utf-8")) > MAX_NAME_BYTES
        or value in {".", ".."}
        or ".." in value
        or value.endswith(".")
        or any(ord(char) < 32 or ord(char) == 127 for char in value)
        or _INVALID_NAME_CHARS.intersection(value)
    ):
        raise M365Error(
            "invalid_payload",
            f'{field} must be a single file or folder name without " * : < > ? / \\ | or ..',
        )
    return value


def folder_path(value: Any, field: str) -> str:
    """Return a normalised absolute path: ``/`` or ``/A/B``."""
    if not isinstance(value, str) or len(value) > MAX_PATH_CHARS:
        raise M365Error("invalid_payload", f"{field} must be a folder path like /Invoices/2026")
    stripped = value.strip()
    if stripped in {"", "/"}:
        return "/"
    segments = [segment for segment in stripped.strip("/").split("/")]
    return "/" + "/".join(name(segment, field) for segment in segments)


def item_id(value: Any, field: str) -> str:
    if not isinstance(value, str) or not _ITEM_ID.match(value) or ".." in value:
        raise M365Error("invalid_payload", f"{field} is not a valid drive item ID")
    return value


def drive_id(value: Any, field: str) -> str:
    if not isinstance(value, str) or not _DRIVE_ID.match(value):
        raise M365Error("invalid_payload", f"{field} is not a valid drive ID")
    return value


def file_types(value: Any) -> list[str]:
    if not isinstance(value, list) or not 1 <= len(value) <= 20:
        raise M365Error("invalid_payload", "file_types must list 1 to 20 file extensions")
    result: list[str] = []
    for raw in value:
        extension = raw.strip().lower().lstrip(".") if isinstance(raw, str) else ""
        if not _EXTENSION.match(extension):
            raise M365Error("invalid_payload", "file_types must hold extensions like pdf or xlsx")
        if extension not in result:
            result.append(extension)
    return result


def bounded_int(value: Any, *, field: str, minimum: int, maximum: int, default: int) -> int:
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise M365Error(
            "invalid_payload", f"{field} must be an integer from {minimum} to {maximum}"
        )
    return value


def choice(value: Any, *, field: str, allowed: tuple[str, ...], default: str) -> str:
    if value is None:
        return default
    if not isinstance(value, str) or value not in allowed:
        raise M365Error("invalid_payload", f"{field} must be one of {', '.join(allowed)}")
    return value


def bounded_text(value: Any, *, field: str, maximum: int, required: bool = True) -> str:
    if value is None and not required:
        return ""
    if (
        not isinstance(value, str)
        or not value.strip()
        or len(value) > maximum
        or any(ord(char) < 32 or ord(char) == 127 for char in value)
    ):
        raise M365Error("invalid_payload", f"{field} must be non-empty text up to {maximum} chars")
    return value.strip()


# -- Graph paths --------------------------------------------------------------


def encode_path(path: str) -> str:
    """``/A B/c`` → ``A%20B/c`` (no leading slash)."""
    return "/".join(quote(segment, safe="") for segment in path.strip("/").split("/") if segment)


def quote_id(value: str) -> str:
    return quote(value, safe="!")


def item_ref(base: str, *, item: str = "", path: str = "") -> str:
    """Graph path of one item addressed by id or by drive path (``/`` is the root)."""
    if item:
        return f"{base}/items/{quote_id(item)}"
    if not path or path == "/":
        return f"{base}/root"
    return f"{base}/root:/{encode_path(path)}:"


def child_ref(base: str, *, parent_item: str = "", parent_path: str = "", child: str) -> str:
    """Graph path of ``child`` inside a folder, ending with ``:`` for path addressing."""
    encoded = quote(child, safe="")
    if parent_item:
        return f"{base}/items/{quote_id(parent_item)}:/{encoded}:"
    if not parent_path or parent_path == "/":
        return f"{base}/root:/{encoded}:"
    return f"{base}/root:/{encode_path(parent_path)}/{encoded}:"


def children_ref(base: str, *, folder_item: str = "", folder_path_value: str = "") -> str:
    return f"{item_ref(base, item=folder_item, path=folder_path_value)}/children"


def item_address(
    raw: dict[str, Any], *, id_field: str, path_field: str, required: bool
) -> tuple[str, str]:
    """Return ``(item_id, path)`` from an input with at most one of the two fields."""
    has_id = raw.get(id_field) not in (None, "")
    has_path = raw.get(path_field) not in (None, "")
    if has_id and has_path:
        raise M365Error("invalid_payload", f"Give either {id_field} or {path_field}, not both")
    if has_id:
        return item_id(raw.get(id_field), id_field), ""
    if has_path:
        return "", folder_path(raw.get(path_field), path_field)
    if required:
        raise M365Error("invalid_payload", f"{id_field} or {path_field} is required")
    return "", "/"


__all__ = [
    "NOT_FOUND_CODES",
    "bounded_int",
    "bounded_text",
    "child_ref",
    "children_ref",
    "choice",
    "drive_id",
    "encode_path",
    "file_types",
    "folder_path",
    "item_address",
    "item_id",
    "item_missing",
    "item_ref",
    "name",
    "quote_id",
]
