"""Shared fakes for the mail tests: an in-memory Outlook mailbox behind a fake transport.

This module holds no tests. The fake applies Graph effects to in-memory state and can
inject faults per action: ``lost`` (the effect happens, the response is lost),
``dropped`` (no effect, ambiguous), ``late`` (no effect now, applied by
``land_late()``), ``lost_twice`` (two effects, response lost), ``duplicated`` (the effect happens twice, the response is normal),
``phantom`` (a normal response, but no effect), ``refused`` (a confirmed 400) and
``throttled`` (a confirmed 429).
"""

from __future__ import annotations

import base64
import hashlib
import json
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import unquote

from m365.errors import M365Error
from m365.transport import GraphResponse
from mail.send import KEY_PROPERTY_ID, STATE_PROPERTY_ID

STEP_EXECUTION_ID = "6a1b2c3d-4e5f-4a6b-8c7d-000000000101"
USER_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-acme-mail",
    "connection_type_id": "m365_user",
    "config": {},
    "secrets": {"access_token": "opaque-token"},
}
APP_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-acme-mail",
    "connection_type_id": "m365_mail_app",
    "config": {
        "tenant_id": "acme.onmicrosoft.com",
        "client_id": "11111111-2222-3333-4444-555555555555",
        "mailbox": "cmr@acme.example",
        "control_mailbox": "ceo@acme.example",
    },
    "secrets": {"client_secret": "app-secret"},
}
ATTACHMENT_OVERHEAD = 180


def action_payload(
    action_id: str,
    operation_input: dict[str, Any] | None = None,
    *,
    connection: dict[str, Any] | None = None,
    test_mode: bool = False,
    step_execution_id: str = STEP_EXECUTION_ID,
    artifacts: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "contract_version": "extension_host_v2",
        "mode": "action",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": "send_reply",
            "step_execution_id": step_execution_id,
            "test_mode": test_mode,
        },
        "action": {
            "action_id": action_id,
            "input": {"connection_ref": "conn-acme-mail", **(operation_input or {})},
            "target": {"connection": dict(connection or USER_CONNECTION)},
        },
    }
    if artifacts is not None:
        payload["artifacts"] = artifacts
    return payload


def input_artifacts(files: dict[str, tuple[str, str, bytes]]) -> dict[str, Any]:
    """Descriptors for ``{handle: (filename, content_type, bytes)}``."""
    return {
        "inputs": [
            {
                "artifact_id": handle,
                "artifact_handle": handle,
                "binding_key": "attachments",
                "filename": name,
                "content_type": content_type,
                "size_bytes": len(data),
                "access": {
                    "transport": "presigned_url",
                    "mode": "read",
                    "max_size_bytes": 200 * 1024 * 1024,
                },
            }
            for handle, (name, content_type, data) in files.items()
        ]
    }


def streamer_for(files: dict[str, tuple[str, str, bytes]]) -> Callable[..., Any]:
    def stream(_payload: Any, *, artifact_id: str, chunk_size: int, timeout_seconds: float):
        data = files[artifact_id][2]
        for offset in range(0, len(data), chunk_size):
            yield data[offset : offset + chunk_size]

    return stream


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass
class Call:
    method: str
    url: str
    query: dict[str, Any]
    json_body: Any
    raw_body: bytes | None
    headers: dict[str, str]
    retry: str
    authorized: bool


@dataclass
class Message:
    id: str
    folder: str
    is_draft: bool
    created: str
    props: dict[str, str] = field(default_factory=dict)
    attachments: list[dict[str, Any]] = field(default_factory=list)
    is_read: bool = False
    flag: str = "notFlagged"
    subject: str = ""
    body: dict[str, Any] = field(default_factory=dict)
    to: list[Any] = field(default_factory=list)


def ambiguous(what: str) -> M365Error:
    return M365Error("timeout_unknown", f"{what} did not complete", http_status=504, ambiguous=True)


class FakeOutlook:
    """An in-memory mailbox reachable at ``/me``."""

    FOLDERS = {
        "inbox": "FOLDER-INBOX",
        "drafts": "FOLDER-DRAFTS",
        "sentitems": "FOLDER-SENT",
        "deleteditems": "FOLDER-DELETED",
        "archive": "FOLDER-ARCHIVE",
        "junkemail": "FOLDER-JUNK",
    }

    def __init__(self, base: str = "/me") -> None:
        self.base = base
        self.messages: dict[str, Message] = {}
        self.calls: list[Call] = []
        self.faults: dict[str, list[str]] = {}
        self.pending: list[Callable[[], None]] = []
        self.sessions: dict[str, dict[str, Any]] = {}
        self.counter = 0
        self.sent_log: list[str] = []
        #: After a send, GET by the draft id still shows a draft (Exchange is slow).
        self.send_invisible = False
        #: After a send, the sent copy has a new id (as without immutable ids).
        self.ids_change_on_send = False
        #: Called with every request before it is handled.
        self.hook: Callable[[Call], None] | None = None

    # -- scripting ------------------------------------------------------------------

    def fault(self, op: str, *modes: str) -> None:
        self.faults.setdefault(op, []).extend(modes)

    def land_late(self) -> None:
        pending, self.pending = self.pending, []
        for effect in pending:
            effect()

    def add_message(self, folder: str = "inbox", **fields: Any) -> Message:
        message = Message(
            id=self._id("MSG"),
            folder=self.FOLDERS.get(folder, folder),
            is_draft=False,
            created=self._now(),
            **fields,
        )
        self.messages[message.id] = message
        return message

    def drafts(self) -> list[Message]:
        return [m for m in self.messages.values() if m.folder == self.FOLDERS["drafts"]]

    def sent(self) -> list[Message]:
        return [m for m in self.messages.values() if m.folder == self.FOLDERS["sentitems"]]

    def writes(self) -> list[Call]:
        return [call for call in self.calls if call.method != "GET"]

    def remaining_seconds(self) -> float:
        return 200.0

    # -- transport --------------------------------------------------------------------

    def request(
        self,
        method: str,
        path_or_url: str,
        *,
        query: dict[str, Any] | None = None,
        json_body: Any = None,
        raw_body: bytes | None = None,
        content_type: str | None = None,
        headers: dict[str, str] | None = None,
        response_limit: int = 4 * 1024 * 1024,
        retry: str,
        purpose: str,
        authorized: bool = True,
        ok_statuses: Any = (200, 201, 202, 204),
    ) -> GraphResponse:
        assert retry in {"read", "idempotent", "none"}
        if method == "GET":
            assert retry == "read"
        call = Call(
            method,
            path_or_url,
            dict(query or {}),
            json_body,
            raw_body,
            dict(headers or {}),
            retry,
            authorized,
        )
        self.calls.append(call)
        if self.hook is not None:
            self.hook(call)
        if path_or_url.startswith("https://outlook.office.com/upload/"):
            assert not authorized
            return self._upload_chunk(path_or_url, raw_body or b"", dict(headers or {}))
        assert authorized
        path = path_or_url
        if path == "/me" and method == "GET":
            return _json(
                {
                    "displayName": "Ana Acme",
                    "mail": "ana@acme.example",
                    "userPrincipalName": "ana@acme.example",
                }
            )
        assert path.startswith(self.base + "/"), path
        rest = path[len(self.base) :]
        for pattern, handler in self._routes():
            match = re.fullmatch(pattern, rest)
            if match and handler[0] == method:
                return handler[1](
                    *[unquote(group) for group in match.groups()],
                    query=query or {},
                    body=json_body,
                    limit=response_limit,
                )
        raise AssertionError(f"unrouted {method} {path}")

    def _routes(self) -> list[tuple[str, tuple[str, Callable[..., GraphResponse]]]]:
        return [
            (r"/mailFolders/([^/]+)/messages", ("GET", self._list_folder)),
            (r"/mailFolders/([^/]+)", ("GET", self._get_folder)),
            (r"/messages", ("POST", self._create)),
            (
                r"/messages/([^/]+)/(createReply|createReplyAll|createForward)",
                ("POST", self._create_from),
            ),
            (r"/messages/([^/]+)/attachments/createUploadSession", ("POST", self._session)),
            (r"/messages/([^/]+)/attachments", ("GET", self._list_attachments)),
            (r"/messages/([^/]+)/attachments", ("POST", self._add_attachment)),
            (r"/messages/([^/]+)/attachments/([^/]+)/\$value", ("GET", self._attachment_value)),
            (r"/messages/([^/]+)/attachments/([^/]+)", ("GET", self._get_attachment)),
            (r"/messages/([^/]+)/attachments/([^/]+)", ("DELETE", self._delete_attachment)),
            (r"/messages/([^/]+)/send", ("POST", self._send)),
            (r"/messages/([^/]+)/move", ("POST", self._move)),
            (r"/messages/([^/]+)", ("GET", self._get_message)),
            (r"/messages/([^/]+)", ("PATCH", self._patch)),
            (r"/messages/([^/]+)", ("DELETE", self._delete_message)),
        ]

    # -- helpers ----------------------------------------------------------------------

    def _id(self, prefix: str) -> str:
        self.counter += 1
        return f"AAMk-{prefix}-{self.counter:04d}"

    def _now(self) -> str:
        return f"2026-11-22T10:{self.counter // 60:02d}:{self.counter % 60:02d}Z"

    def _message(self, message_id: str) -> Message:
        message = self.messages.get(message_id)
        if message is None:
            raise M365Error("item_not_found", "not found", http_status=404)
        return message

    def _apply(self, op: str, effect: Callable[[], GraphResponse]) -> GraphResponse:
        mode = (self.faults.get(op) or [None]).pop(0) if self.faults.get(op) else None
        if mode is None:
            return effect()
        if mode == "lost":
            effect()
            raise ambiguous(op)
        if mode == "dropped":
            raise ambiguous(op)
        if mode == "late":
            self.pending.append(effect)
            raise ambiguous(op)
        if mode == "lost_twice":
            effect()
            effect()
            raise ambiguous(op)
        if mode == "duplicated":
            effect()
            return effect()
        if mode == "phantom":
            return _json({"id": "AAMk-PHANTOM"}, 201)
        if mode == "refused":
            raise M365Error("invalid_payload", f"{op} refused", http_status=400)
        if mode == "throttled":
            raise M365Error("rate_limited", f"{op} throttled", http_status=429)
        raise AssertionError(mode)

    # -- folders and listing ----------------------------------------------------------

    def _folder_id(self, ref: str) -> str:
        if ref in self.FOLDERS:
            return self.FOLDERS[ref]
        if ref in self.FOLDERS.values():
            return ref
        raise M365Error("item_not_found", "folder", http_status=404)

    def _get_folder(self, ref: str, **_: Any) -> GraphResponse:
        folder_id = self._folder_id(ref)
        inside = [m for m in self.messages.values() if m.folder == folder_id]
        return _json(
            {
                "id": folder_id,
                "displayName": ref,
                "totalItemCount": len(inside),
                "unreadItemCount": sum(1 for m in inside if not m.is_read),
            }
        )

    def _list_folder(self, ref: str, *, query: dict[str, Any], **_: Any) -> GraphResponse:
        folder_id = self._folder_id(ref)
        rows = [m for m in self.messages.values() if m.folder == folder_id]
        flt = str(query.get("$filter") or "")
        match = re.search(r"ep/value eq '((?:[^']|'')*)'", flt)
        if match:
            assert KEY_PROPERTY_ID in flt
            key = match.group(1).replace("''", "'")
            rows = [m for m in rows if m.props.get(KEY_PROPERTY_ID) == key]
        expand = str(query.get("$expand") or "")
        values = []
        for m in rows:
            row: dict[str, Any] = {
                "id": m.id,
                "createdDateTime": m.created,
                "isDraft": m.is_draft,
                "parentFolderId": m.folder,
                "subject": m.subject,
            }
            if expand and STATE_PROPERTY_ID in expand and STATE_PROPERTY_ID in m.props:
                row["singleValueExtendedProperties"] = [
                    {"id": STATE_PROPERTY_ID, "value": m.props[STATE_PROPERTY_ID]}
                ]
            values.append(row)
        return _json({"value": values})

    # -- drafts -------------------------------------------------------------------------

    def _new_draft(
        self, body: dict[str, Any], *, attachments: list[dict[str, Any]] | None = None
    ) -> Message:
        props = {row["id"]: row["value"] for row in body.get("singleValueExtendedProperties") or []}
        for header in body.get("internetMessageHeaders") or []:
            props[f"header:{header['name']}"] = header["value"]
        draft = Message(
            id=self._id("DRAFT"),
            folder=self.FOLDERS["drafts"],
            is_draft=True,
            created=self._now(),
            props=props,
            subject=str(body.get("subject") or ""),
            body=dict(body.get("body") or {}),
            to=list(body.get("toRecipients") or []),
            attachments=[dict(a) for a in attachments or []],
        )
        self.messages[draft.id] = draft
        return draft

    def _create(self, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        return self._apply("create", lambda: _json({"id": self._new_draft(body).id}, 201))

    def _create_from(
        self, original_id: str, action: str, *, body: dict[str, Any], **_: Any
    ) -> GraphResponse:
        original = self._message(original_id)
        message = dict(body.get("message") or {})
        copied = (
            [dict(a, id=self._id("ATT")) for a in original.attachments]
            if action == "createForward"
            else []
        )

        def effect() -> GraphResponse:
            draft = self._new_draft(message, attachments=copied)
            draft.props["comment"] = str(body.get("comment") or "")
            draft.props["action"] = action
            return _json({"id": draft.id}, 201)

        return self._apply("create", effect)

    def _patch(self, message_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        message = self._message(message_id)

        def effect() -> GraphResponse:
            for row in body.get("singleValueExtendedProperties") or []:
                message.props[row["id"]] = row["value"]
            if "isRead" in body:
                message.is_read = body["isRead"]
            if "flag" in body:
                message.flag = body["flag"]["flagStatus"]
            return _json({"id": message.id})

        return self._apply("patch", effect)

    def _delete_message(self, message_id: str, **_: Any) -> GraphResponse:
        self._message(message_id)

        def effect() -> GraphResponse:
            self.messages.pop(message_id, None)
            return GraphResponse(204, {}, b"")

        return self._apply("delete_message", effect)

    def _get_message(self, message_id: str, **_: Any) -> GraphResponse:
        m = self._message(message_id)
        return _json(
            {
                "id": m.id,
                "isDraft": m.is_draft,
                "parentFolderId": m.folder,
                "subject": m.subject,
                "isRead": m.is_read,
                "flag": {"flagStatus": m.flag},
            }
        )

    def _send(self, message_id: str, **_: Any) -> GraphResponse:
        draft = self._message(message_id)
        assert draft.is_draft, "a sent message was sent again"
        assert draft.props.get(STATE_PROPERTY_ID) != "abandoned", "an abandoned draft was sent"

        def effect() -> GraphResponse:
            self.sent_log.append(draft.id)
            if not self.send_invisible:
                draft.is_draft = False
                draft.folder = self.FOLDERS["sentitems"]
            if self.ids_change_on_send:
                self.messages.pop(draft.id)
                draft.id = self._id("SENT")
                self.messages[draft.id] = draft
            return GraphResponse(202, {}, b"")

        return self._apply("send", effect)

    def _move(self, message_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        message = self._message(message_id)
        destination = self._folder_id(body["destinationId"])

        def effect() -> GraphResponse:
            message.folder = destination
            return _json({"id": message.id}, 201)

        return self._apply("move", effect)

    # -- attachments --------------------------------------------------------------------

    def _attachment_row(self, row: dict[str, Any]) -> dict[str, Any]:
        return {
            "@odata.type": f"#microsoft.graph.{row.get('kind', 'file')}Attachment",
            "id": row["id"],
            "name": row["name"],
            "contentType": row["contentType"],
            "size": len(row["content"]) + ATTACHMENT_OVERHEAD,
            "isInline": False,
        }

    def _list_attachments(self, message_id: str, **_: Any) -> GraphResponse:
        message = self._message(message_id)
        return _json({"value": [self._attachment_row(row) for row in message.attachments]})

    def _get_attachment(self, message_id: str, attachment_id: str, **_: Any) -> GraphResponse:
        return _json(self._attachment_row(self._find_attachment(message_id, attachment_id)))

    def _find_attachment(self, message_id: str, attachment_id: str) -> dict[str, Any]:
        for row in self._message(message_id).attachments:
            if row["id"] == attachment_id:
                return row
        raise M365Error("item_not_found", "attachment", http_status=404)

    def _attachment_value(
        self, message_id: str, attachment_id: str, *, limit: int, **_: Any
    ) -> GraphResponse:
        content = self._find_attachment(message_id, attachment_id)["content"]
        if len(content) > limit:
            raise M365Error("response_too_large", "too large")
        return GraphResponse(200, {}, content)

    def attach(
        self, message: Message, name: str, content_type: str, content: bytes, kind: str = "file"
    ) -> str:
        attachment_id = self._id("ATT")
        message.attachments.append(
            {
                "id": attachment_id,
                "name": name,
                "contentType": content_type,
                "content": content,
                "kind": kind,
            }
        )
        return attachment_id

    def _add_attachment(self, message_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        message = self._message(message_id)
        assert body["@odata.type"] == "#microsoft.graph.fileAttachment"
        content = base64.b64decode(body["contentBytes"])

        def effect() -> GraphResponse:
            if message_id in self.messages:  # a late add on a deleted draft fails harmlessly
                return _json(
                    {"id": self.attach(message, body["name"], body["contentType"], content)}, 201
                )
            return GraphResponse(404, {}, b"")

        return self._apply("attach", effect)

    def _delete_attachment(self, message_id: str, attachment_id: str, **_: Any) -> GraphResponse:
        message = self._message(message_id)
        self._find_attachment(message_id, attachment_id)
        message.attachments = [row for row in message.attachments if row["id"] != attachment_id]
        return GraphResponse(204, {}, b"")

    def _session(self, message_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        self._message(message_id)
        item = body["AttachmentItem"]

        def effect() -> GraphResponse:
            session_id = self._id("SESSION")
            self.sessions[session_id] = {"message": message_id, "item": item, "buffer": bytearray()}
            return _json(
                {
                    "uploadUrl": f"https://outlook.office.com/upload/{session_id}?authtoken=secret-upload-token",
                    "expirationDateTime": "2026-11-22T12:00:00Z",
                },
                201,
            )

        return self._apply("session", effect)

    def _upload_chunk(self, url: str, chunk: bytes, headers: dict[str, str]) -> GraphResponse:
        assert "Authorization" not in headers
        session_id = url.split("/upload/", 1)[1].split("?", 1)[0]
        session = self.sessions[session_id]
        start, end, total = map(
            int, re.fullmatch(r"bytes (\d+)-(\d+)/(\d+)", headers["Content-Range"]).groups()
        )
        assert (
            start == len(session["buffer"])
            and end - start + 1 == len(chunk)
            and len(chunk) <= 4 * 1024 * 1024
        )
        final = end + 1 == total

        def effect() -> GraphResponse:
            session["buffer"].extend(chunk)
            if not final:
                return _json({"nextExpectedRanges": [f"{end + 1}-"]})
            message = self.messages.get(session["message"])
            if message is None:
                return GraphResponse(404, {}, b"")
            item = session["item"]
            self.attach(message, item["name"], item["contentType"], bytes(session["buffer"]))
            return GraphResponse(201, {"location": "https://outlook.office.com/attachment"}, b"")

        return self._apply("chunk_final" if final else "chunk", effect)


class ScriptedGraph:
    """Answers requests in order from a script of dicts, bytes or exceptions; records calls."""

    def __init__(self, *script: Any) -> None:
        self.script = list(script)
        self.calls: list[Call] = []

    def remaining_seconds(self) -> float:
        return 200.0

    def request(self, method: str, path_or_url: str, **kwargs: Any) -> GraphResponse:
        self.calls.append(
            Call(
                method,
                path_or_url,
                dict(kwargs.get("query") or {}),
                kwargs.get("json_body"),
                kwargs.get("raw_body"),
                dict(kwargs.get("headers") or {}),
                kwargs["retry"],
                kwargs.get("authorized", True),
            )
        )
        if not self.script:
            raise AssertionError(f"unexpected {method} {path_or_url}")
        step = self.script.pop(0)
        if isinstance(step, Exception):
            raise step
        if isinstance(step, bytes):
            if len(step) > kwargs.get("response_limit", 4 * 1024 * 1024):
                raise M365Error("response_too_large", "too large")
            return GraphResponse(200, {}, step)
        return _json(step)


def _json(value: Any, status: int = 200) -> GraphResponse:
    return GraphResponse(status, {}, json.dumps(value).encode("utf-8"))


__all__ = [
    "APP_CONNECTION",
    "STEP_EXECUTION_ID",
    "USER_CONNECTION",
    "FakeOutlook",
    "ScriptedGraph",
    "action_payload",
    "input_artifacts",
    "sha",
    "streamer_for",
]
