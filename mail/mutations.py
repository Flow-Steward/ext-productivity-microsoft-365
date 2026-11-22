"""Idempotent mailbox writes: flags, move and move-to-Deleted-Items.

These three are on the B6 retry allow-list: a repeat finds the message already in the
requested state, so the transport may retry them (``retry="idempotent"``).
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import (
    Mailbox,
    bool_input,
    folder_ref,
    message_ids_input,
    not_found_as,
    provider_id,
    segment,
)

FLAG_STATUSES = frozenset({"flagged", "complete", "notFlagged"})
_MESSAGE_NOT_FOUND = "The message was not found in this mailbox"


def set_message_flags(ctx: OperationContext) -> dict[str, Any]:
    ids = message_ids_input(ctx.input.get("message_ids"))
    is_read = bool_input(ctx.input.get("is_read"), "is_read")
    flag_status = ctx.input.get("flag_status")
    if flag_status is not None and flag_status not in FLAG_STATUSES:
        raise M365Error("invalid_payload", "flag_status must be flagged, complete or notFlagged")
    if is_read is None and flag_status is None:
        raise M365Error("invalid_payload", "Set is_read, flag_status or both")
    patch: dict[str, Any] = {}
    if is_read is not None:
        patch["isRead"] = is_read
    if flag_status is not None:
        patch["flag"] = {"flagStatus": flag_status}
    client = Mailbox.for_connection(ctx.transport, ctx.connection)

    def apply(message_id: str) -> dict[str, Any]:
        response = client.write(
            "PATCH",
            f"/messages/{segment(message_id)}",
            json_body=patch,
            retry="idempotent",
            purpose="Outlook message flags",
        )
        new_id = _returned_id(response, message_id)
        return {"message_id": new_id, "status": "succeeded", "changed": True}

    return _per_message(ids, apply)


def move_messages(ctx: OperationContext) -> dict[str, Any]:
    ids = message_ids_input(ctx.input.get("message_ids"))
    destination = folder_ref(ctx.input.get("destination_folder"), "destination_folder")
    client = Mailbox.for_connection(ctx.transport, ctx.connection)
    folder_id = _folder_id(client, destination)
    return _per_message(ids, lambda message_id: _move(client, message_id, folder_id))


def delete_messages(ctx: OperationContext) -> dict[str, Any]:
    """Move to Deleted Items only; there is no permanent delete."""
    ids = message_ids_input(ctx.input.get("message_ids"))
    client = Mailbox.for_connection(ctx.transport, ctx.connection)
    folder_id = _folder_id(client, "deleteditems")
    return _per_message(ids, lambda message_id: _move(client, message_id, folder_id))


def _folder_id(client: Mailbox, folder: str) -> str:
    with not_found_as("folder_not_found", "The destination folder was not found in this mailbox"):
        row = client.get(
            f"/mailFolders/{segment(folder)}",
            query={"$select": "id"},
            response_limit=64 * 1024,
            purpose="Outlook folder",
        )
    return provider_id(row.get("id"))


def _move(client: Mailbox, message_id: str, folder_id: str) -> dict[str, Any]:
    current = client.get(
        f"/messages/{segment(message_id)}",
        query={"$select": "id,parentFolderId"},
        response_limit=64 * 1024,
        purpose="Outlook message location",
    )
    current_id = str(current.get("id") or message_id)
    if str(current.get("parentFolderId") or "") == folder_id:
        return {"message_id": current_id, "status": "succeeded", "changed": False}
    response = client.write(
        "POST",
        f"/messages/{segment(message_id)}/move",
        json_body={"destinationId": folder_id},
        retry="idempotent",
        purpose="Outlook message move",
    )
    return {
        "message_id": _returned_id(response, message_id),
        "status": "succeeded",
        "changed": True,
    }


def _returned_id(response: Any, fallback: str) -> str:
    try:
        value = response.json().get("id")
    except M365Error:
        return fallback
    return value if isinstance(value, str) and value else fallback


def _per_message(ids: list[str], apply: Callable[[str], dict[str, Any]]) -> dict[str, Any]:
    """Apply to each message; stop at an unconfirmed write, never hide a partial effect."""
    results: list[dict[str, Any]] = []
    first_error: M365Error | None = None
    for index, message_id in enumerate(ids):
        try:
            with not_found_as("message_not_found", _MESSAGE_NOT_FOUND):
                row = apply(message_id)
            row["requested_message_id"] = message_id
            results.append(row)
        except M365Error as exc:
            if exc.ambiguous:
                results.append(
                    {
                        "requested_message_id": message_id,
                        "message_id": message_id,
                        "status": "timeout_unknown",
                        "error_code": "timeout_unknown",
                    }
                )
                results.extend(
                    {
                        "requested_message_id": rest,
                        "message_id": rest,
                        "status": "not_attempted",
                        "changed": False,
                    }
                    for rest in ids[index + 1 :]
                )
                break
            first_error = first_error or exc
            results.append(
                {
                    "requested_message_id": message_id,
                    "message_id": message_id,
                    "status": "failed",
                    "error_code": exc.code,
                    "changed": False,
                }
            )
            if exc.code != "message_not_found":
                # Access, throttling or mailbox problems affect the rest too.
                results.extend(
                    {
                        "requested_message_id": rest,
                        "message_id": rest,
                        "status": "not_attempted",
                        "changed": False,
                    }
                    for rest in ids[index + 1 :]
                )
                break
    statuses = {row["status"] for row in results}
    if first_error is not None and statuses <= {"failed", "not_attempted"}:
        # Nothing changed anywhere: report the failure itself.
        raise first_error
    if "timeout_unknown" in statuses or ("succeeded" in statuses and statuses - {"succeeded"}):
        aggregate = "timeout_unknown"
    else:
        aggregate = "succeeded"
    changed = sum(1 for row in results if row.get("changed"))
    return {
        "results": results,
        "changed_count": changed,
        "changed": changed > 0,
        "external_effect_status": aggregate,
        "definitely_no_external_effect": False,
    }


__all__ = ["FLAG_STATUSES", "delete_messages", "move_messages", "set_message_flags"]
