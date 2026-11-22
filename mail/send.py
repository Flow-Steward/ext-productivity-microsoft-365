"""Draft-then-send for ``send_message``, ``reply_message`` and ``forward_message`` (task B5.1).

1. Create a draft that carries the step's idempotency key twice: as the single-value
   extended property ``FlowStewardIdempotencyKey`` (searchable with ``$filter``) and as
   the internet header ``X-FlowSteward-Idempotency-Key``. An ambiguous create is never
   repeated blindly: Drafts is searched by the key first.
2. Add the attachments (inline below 3 MB, otherwise an attachment upload session),
   then reconcile the draft's attachments with the expected set by SHA-256 and count.
   If an add stayed ambiguous, the draft is tainted for good: it is marked
   ``FlowStewardDraftState=abandoned``, deleted if possible, and the step fails with
   ``attachment_state_unknown`` and the provider retry signal.
3. Send the draft. An ambiguous send is confirmed through the draft's immutable id
   (``isDraft: false``) or, as a fallback, a Sent Items search by the key; otherwise the
   result is ``timeout_unknown``. Nothing here is ever re-sent automatically.

Every write in this module uses ``retry="none"``: the transport never repeats it.
"""

from __future__ import annotations

import base64
import hashlib
import time
from collections import Counter
from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass
from typing import Any

from flowsteward_extension_sdk import report_progress

from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import (
    Mailbox,
    graph_id,
    not_found_as,
    odata_string,
    provider_id,
    recipients_input,
    segment,
    text_input,
    without_ambiguity,
)
from .outgoing import (
    OutgoingAttachment,
    attachment_rows,
    load_attachments,
    normalized_content_type,
    read_chunks,
)

#: Fixed Flow Steward property set for the extended properties below.
PROPERTY_SET_GUID = "b9fd5293-1be4-448e-b08f-b4467c228ad8"
KEY_PROPERTY_ID = f"String {{{PROPERTY_SET_GUID}}} Name FlowStewardIdempotencyKey"
STATE_PROPERTY_ID = f"String {{{PROPERTY_SET_GUID}}} Name FlowStewardDraftState"
HEADER_NAME = "X-FlowSteward-Idempotency-Key"
#: PS_INTERNET_HEADERS: a named string property in this set is sent as an internet header.
#: Used for replies and forwards, whose create actions take no ``internetMessageHeaders``.
HEADER_PROPERTY_ID = f"String {{00020386-0000-0000-C000-000000000046}} Name {HEADER_NAME}"
ABANDONED = "abandoned"
#: Each upload-session request stays below Outlook's 4 MB per request.
UPLOAD_CHUNK_BYTES = 10 * 320 * 1024
MAX_BODY_CHARS = 1_000_000
MAX_SUBJECT_CHARS = 998
DRAFT_LOOKUP_ATTEMPTS = 3
DRAFT_LOOKUP_DELAY_SECONDS = 2.0
CONFIRM_BUDGET_SECONDS = 90.0
CONFIRM_INTERVAL_SECONDS = 5.0
RECONCILE_ROUNDS = 3

#: Indirection points for tests.
sleep: Callable[[float], None] = time.sleep
clock: Callable[[], float] = time.monotonic

SEND_FIELDS = frozenset(
    {
        "to",
        "cc",
        "bcc",
        "reply_to",
        "subject",
        "body_text",
        "body_html",
        "importance",
        "attachments",
    }
)
REPLY_FIELDS = frozenset({"message_id", "reply_all", "comment", "attachments"})
FORWARD_FIELDS = frozenset({"message_id", "to", "cc", "bcc", "comment", "attachments"})


class _Tainted(Exception):
    """An attachment add stayed ambiguous: this draft must never be sent."""


@dataclass(frozen=True)
class DraftPlan:
    kind: str
    create_path: str
    create_body: dict[str, Any]


@dataclass(frozen=True)
class FoundDraft:
    message_id: str
    created: str
    state: str


# -- input --------------------------------------------------------------------------


def _key_properties(key: str, *, header_as_property: bool) -> list[dict[str, str]]:
    rows = [{"id": KEY_PROPERTY_ID, "value": key}]
    if header_as_property:
        rows.append({"id": HEADER_PROPERTY_ID, "value": key})
    return rows


def _comment(raw: Mapping[str, Any]) -> str:
    return text_input(
        raw.get("comment"), "comment", max_length=MAX_BODY_CHARS, allow_newlines=True, strip=False
    )


def plan_send(raw: Mapping[str, Any], key: str) -> DraftPlan:
    to = recipients_input(raw.get("to"), "to")
    cc = recipients_input(raw.get("cc"), "cc")
    bcc = recipients_input(raw.get("bcc"), "bcc")
    reply_to = recipients_input(raw.get("reply_to"), "reply_to", maximum=10)
    if not (to or cc or bcc):
        raise M365Error("invalid_payload", "At least one recipient in to, cc or bcc is required")
    if len(to) + len(cc) + len(bcc) > 500:
        raise M365Error("invalid_payload", "A message can have at most 500 recipients")
    body_text = text_input(
        raw.get("body_text"),
        "body_text",
        max_length=MAX_BODY_CHARS,
        allow_newlines=True,
        strip=False,
    )
    body_html = text_input(
        raw.get("body_html"),
        "body_html",
        max_length=MAX_BODY_CHARS,
        allow_newlines=True,
        strip=False,
    )
    if body_text and body_html:
        raise M365Error("invalid_payload", "Use body_text or body_html, not both")
    importance = raw.get("importance") or "normal"
    if importance not in {"low", "normal", "high"}:
        raise M365Error("invalid_payload", "importance must be low, normal or high")
    message: dict[str, Any] = {
        "subject": text_input(raw.get("subject"), "subject", max_length=MAX_SUBJECT_CHARS),
        "body": {"contentType": "HTML" if body_html else "Text", "content": body_html or body_text},
        "toRecipients": to,
        "ccRecipients": cc,
        "bccRecipients": bcc,
        "importance": importance,
        "singleValueExtendedProperties": _key_properties(key, header_as_property=False),
        "internetMessageHeaders": [{"name": HEADER_NAME, "value": key}],
    }
    if reply_to:
        message["replyTo"] = reply_to
    return DraftPlan("send", "/messages", message)


def plan_reply(raw: Mapping[str, Any], key: str) -> DraftPlan:
    original = graph_id(raw.get("message_id"), "message_id")
    reply_all = raw.get("reply_all", False)
    if not isinstance(reply_all, bool):
        raise M365Error("invalid_payload", "reply_all must be true or false")
    action = "createReplyAll" if reply_all else "createReply"
    body: dict[str, Any] = {
        "message": {"singleValueExtendedProperties": _key_properties(key, header_as_property=True)}
    }
    comment = _comment(raw)
    if comment:
        body["comment"] = comment
    return DraftPlan("reply", f"/messages/{segment(original)}/{action}", body)


def plan_forward(raw: Mapping[str, Any], key: str) -> DraftPlan:
    original = graph_id(raw.get("message_id"), "message_id")
    to = recipients_input(raw.get("to"), "to")
    if not to:
        raise M365Error("invalid_payload", "to is required for a forward")
    cc = recipients_input(raw.get("cc"), "cc")
    bcc = recipients_input(raw.get("bcc"), "bcc")
    if len(to) + len(cc) + len(bcc) > 500:
        raise M365Error("invalid_payload", "A message can have at most 500 recipients")
    message: dict[str, Any] = {
        "toRecipients": to,
        "singleValueExtendedProperties": _key_properties(key, header_as_property=True),
    }
    if cc:
        message["ccRecipients"] = cc
    if bcc:
        message["bccRecipients"] = bcc
    body: dict[str, Any] = {"message": message}
    comment = _comment(raw)
    if comment:
        body["comment"] = comment
    return DraftPlan("forward", f"/messages/{segment(original)}/createForward", body)


# -- operations ----------------------------------------------------------------------


def send_message(ctx: OperationContext) -> dict[str, Any]:
    return _run(ctx, plan_send)


def reply_message(ctx: OperationContext) -> dict[str, Any]:
    return _run(ctx, plan_reply)


def forward_message(ctx: OperationContext) -> dict[str, Any]:
    return _run(ctx, plan_forward)


def _run(
    ctx: OperationContext, planner: Callable[[Mapping[str, Any], str], DraftPlan]
) -> dict[str, Any]:
    key = ctx.idempotency_key
    plan = planner(ctx.input, key)
    rows = attachment_rows(ctx.input.get("attachments"))
    attachments = load_attachments(ctx.payload, rows)
    flow = SendFlow(ctx.payload, Mailbox.for_connection(ctx.transport, ctx.connection), key)
    return flow.execute(plan, attachments)


# -- the flow ------------------------------------------------------------------------


class SendFlow:
    def __init__(self, payload: dict[str, Any], client: Mailbox, key: str) -> None:
        self.payload = payload
        self.client = client
        self.key = key
        self.attachment_count = 0

    # public ----------------------------------------------------------------------

    def execute(self, plan: DraftPlan, attachments: list[OutgoingAttachment]) -> dict[str, Any]:
        self.attachment_count = len(attachments)
        try:
            sent = self.find_sent()
            if sent:
                # A sent copy with this key exists: the email already went out.
                return self._success(sent[0], "already_sent")
            self.discard_previous_drafts()
            draft_id = self.create_draft(plan)
            try:
                self.add_attachments(draft_id, attachments)
            except _Tainted:
                raise self.abandon(draft_id) from None
            except M365Error:
                self.cleanup_quietly()
                raise
            self.discard_other_drafts(draft_id)
        except M365Error as exc:
            # Nothing has been sent before step 3, whatever the request outcome was.
            raise without_ambiguity(exc) if exc.ambiguous else exc
        return self.send(draft_id)

    # step 0: resume ------------------------------------------------------------------

    def discard_previous_drafts(self) -> None:
        """A retried step deletes every draft an earlier attempt left and builds a fresh one."""
        for draft in self.find_drafts():
            if not self.delete_message(draft.message_id):
                raise M365Error(
                    "attachment_state_unknown",
                    f"An earlier draft ({draft.message_id}) could not be deleted; nothing was sent",
                    retryable=True,
                    failure_class="provider",
                )

    # step 1: create ----------------------------------------------------------------

    def create_draft(self, plan: DraftPlan) -> str:
        for attempt in (1, 2):
            try:
                with not_found_as(
                    "message_not_found", "The original message was not found in this mailbox"
                ):
                    response = self.client.write(
                        "POST",
                        plan.create_path,
                        json_body=plan.create_body,
                        retry="none",
                        purpose="Outlook draft",
                    )
                return provider_id(response.json().get("id"))
            except M365Error as exc:
                if not exc.ambiguous:
                    self.cleanup_quietly()
                    raise
            found = self.lookup_after_ambiguous_create()
            if found:
                return found
            if attempt == 2:
                break
        raise M365Error(
            "provider_unavailable",
            "Outlook did not confirm the draft; nothing was sent",
            retryable=True,
            failure_class="provider",
        )

    def lookup_after_ambiguous_create(self) -> str | None:
        for attempt in range(DRAFT_LOOKUP_ATTEMPTS):
            drafts = [draft for draft in self.find_drafts() if draft.state != ABANDONED]
            if drafts:
                drafts.sort(key=lambda draft: (draft.created, draft.message_id))
                keep, extras = drafts[0], drafts[1:]
                for extra in extras:
                    if not self.delete_message(extra.message_id):
                        raise M365Error(
                            "provider_unavailable",
                            "A duplicate draft could not be deleted; nothing was sent",
                            retryable=True,
                            failure_class="provider",
                        )
                return keep.message_id
            if attempt + 1 < DRAFT_LOOKUP_ATTEMPTS:
                sleep(DRAFT_LOOKUP_DELAY_SECONDS)
        return None

    def discard_other_drafts(self, draft_id: str) -> None:
        """A create that landed late leaves a second draft; remove it before sending."""
        for draft in self.find_drafts():
            if draft.message_id != draft_id and not self.delete_message(draft.message_id):
                raise M365Error(
                    "provider_unavailable",
                    "A duplicate draft could not be deleted; nothing was sent",
                    retryable=True,
                    failure_class="provider",
                )

    # step 2: attachments -------------------------------------------------------------

    def add_attachments(self, draft_id: str, attachments: list[OutgoingAttachment]) -> None:
        if not attachments:
            return
        # A forward copies the original's attachments; they are not ours to reconcile.
        baseline = {row["id"] for row in self.list_attachments(draft_id)}
        unknown: Counter[str] = Counter()
        for index, attachment in enumerate(attachments):
            report_progress(f"Attaching {attachment.filename}", done=index, total=len(attachments))
            if self.add_one(draft_id, attachment) == "unknown":
                unknown[attachment.sha256] += 1
        self.reconcile(draft_id, baseline, attachments, unknown)

    def add_one(self, draft_id: str, attachment: OutgoingAttachment) -> str:
        """Return ``added`` or ``unknown``; raise on a confirmed failure."""
        if attachment.inline:
            body = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment.filename,
                "contentType": attachment.content_type,
                "contentBytes": base64.b64encode(attachment.content or b"").decode("ascii"),
                "isInline": False,
            }
            try:
                self.client.write(
                    "POST",
                    f"/messages/{segment(draft_id)}/attachments",
                    json_body=body,
                    retry="none",
                    purpose="Outlook attachment",
                )
            except M365Error as exc:
                if exc.ambiguous:
                    return "unknown"
                raise
            return "added"
        return self.upload_large(draft_id, attachment)

    def upload_large(self, draft_id: str, attachment: OutgoingAttachment) -> str:
        upload_url = self._upload_session(draft_id, attachment)
        digest_source = _Hashing(
            read_chunks(self.payload, attachment.handle, expected_size=attachment.size)
        )
        for start, chunk in _rechunk(digest_source, UPLOAD_CHUNK_BYTES):
            end = start + len(chunk) - 1
            final = end + 1 == attachment.size
            if final and digest_source.hexdigest() != attachment.sha256:
                raise M365Error(
                    "artifact_input_unavailable",
                    "An attachment artifact changed while it was uploaded",
                )
            report_progress(
                f"Uploading {attachment.filename} (MB)",
                done=(end + 1) // (1024 * 1024),
                total=max(1, attachment.size // (1024 * 1024)),
            )
            try:
                self.client.transport.request(
                    "PUT",
                    upload_url,
                    raw_body=chunk,
                    content_type="application/octet-stream",
                    headers={"Content-Range": f"bytes {start}-{end}/{attachment.size}"},
                    retry="none",
                    purpose="Outlook attachment upload",
                    authorized=False,
                    ok_statuses=(200, 201),
                )
            except M365Error as exc:
                if final and exc.ambiguous:
                    return "unknown"
                # Without its last byte range a session never becomes an attachment.
                raise _upload_failed(exc) from exc
        return "added"

    def _upload_session(self, draft_id: str, attachment: OutgoingAttachment) -> str:
        body = {
            "AttachmentItem": {
                "attachmentType": "file",
                "name": attachment.filename,
                "size": attachment.size,
                "contentType": attachment.content_type,
            }
        }
        last: M365Error | None = None
        for _attempt in (1, 2):
            try:
                response = self.client.write(
                    "POST",
                    f"/messages/{segment(draft_id)}/attachments/createUploadSession",
                    json_body=body,
                    retry="none",
                    purpose="Outlook attachment upload session",
                )
            except M365Error as exc:
                if not exc.ambiguous:
                    raise _upload_failed(exc) from exc
                # A session without uploaded bytes never becomes an attachment: safe to ask again.
                last = exc
                continue
            url = response.json().get("uploadUrl")
            if not isinstance(url, str) or not url.startswith("https://"):
                raise M365Error("attachment_upload_failed", "Outlook returned no upload session")
            return url
        raise _upload_failed(last or M365Error("provider_unavailable", "no session"))

    def reconcile(
        self,
        draft_id: str,
        baseline: set[str],
        expected: list[OutgoingAttachment],
        unknown: Counter[str],
    ) -> None:
        """Make the draft's attachments equal the expected set by content and count."""
        expected_counts = Counter(item.sha256 for item in expected)
        by_hash = {item.sha256: item for item in expected}
        for _round in range(RECONCILE_ROUNDS):
            actual = [row for row in self.list_attachments(draft_id) if row["id"] not in baseline]
            assigned, foreign = self.classify(draft_id, actual, expected, unknown)
            surplus = list(foreign)
            missing: Counter[str] = Counter()
            for sha, count in expected_counts.items():
                ids = assigned.get(sha, [])
                surplus.extend(ids[count:])
                if len(ids) < count:
                    missing[sha] = count - len(ids)
                else:
                    unknown[sha] = 0  # every request for this content has landed
            if not surplus and not missing:
                return
            if any(unknown[sha] for sha in missing):
                raise _Tainted
            for attachment_id in surplus:
                self.delete_attachment(draft_id, attachment_id)
            for sha, count in missing.items():
                for _ in range(count):
                    if self.add_one(draft_id, by_hash[sha]) == "unknown":
                        unknown[sha] += 1
        raise _Tainted

    def classify(
        self,
        draft_id: str,
        actual: list[dict[str, Any]],
        expected: list[OutgoingAttachment],
        unknown: Counter[str],
    ) -> tuple[dict[str, list[str]], list[str]]:
        """Assign each draft attachment to an expected content hash, or mark it foreign.

        Name and content type only narrow the candidates; the content decides whenever
        the narrowing is not unique or an add for that content was ambiguous.
        """
        groups: dict[tuple[str, str], list[OutgoingAttachment]] = {}
        for item in expected:
            groups.setdefault(item.match_key, []).append(item)
        rows_by_key: dict[tuple[str, str], list[dict[str, Any]]] = {}
        foreign: list[str] = []
        for row in actual:
            key = (row["name"], normalized_content_type(row["content_type"]))
            if row["type"] != "file" or key not in groups:
                foreign.append(row["id"])
            else:
                rows_by_key.setdefault(key, []).append(row)
        assigned: dict[str, list[str]] = {}
        for key, rows in rows_by_key.items():
            candidates = groups[key]
            only = candidates[0]
            if (
                len(candidates) == 1
                and len(rows) == 1
                and not unknown[only.sha256]
                and rows[0]["size"] >= only.size
            ):
                assigned.setdefault(only.sha256, []).append(rows[0]["id"])
                continue
            hashes = {item.sha256 for item in candidates}
            limit = max(item.size for item in candidates)
            for row in rows:
                sha = self.attachment_sha256(draft_id, row["id"], limit)
                if sha in hashes:
                    assigned.setdefault(sha, []).append(row["id"])
                else:
                    foreign.append(row["id"])
        return assigned, foreign

    def attachment_sha256(self, draft_id: str, attachment_id: str, limit: int) -> str | None:
        try:
            content = self.client.get_bytes(
                f"/messages/{segment(draft_id)}/attachments/{segment(attachment_id)}/$value",
                response_limit=max(1, limit),
                purpose="Outlook attachment check",
            )
        except M365Error as exc:
            if exc.code == "response_too_large":
                return None  # larger than any expected file: not ours
            raise
        return hashlib.sha256(content).hexdigest()

    def list_attachments(self, draft_id: str) -> list[dict[str, Any]]:
        rows = self.client.list_all(
            f"/messages/{segment(draft_id)}/attachments",
            query={"$select": "id,name,contentType,size,isInline"},
            max_items=250,
            purpose="Outlook draft attachments",
        )
        result = []
        for row in rows:
            size = row.get("size")
            result.append(
                {
                    "id": provider_id(row.get("id")),
                    "name": str(row.get("name") or ""),
                    "content_type": str(row.get("contentType") or ""),
                    "size": size if isinstance(size, int) and not isinstance(size, bool) else 0,
                    "type": {
                        "#microsoft.graph.fileAttachment": "file",
                        "#microsoft.graph.itemAttachment": "item",
                        "#microsoft.graph.referenceAttachment": "reference",
                    }.get(str(row.get("@odata.type") or ""), "unknown"),
                }
            )
        return result

    def delete_attachment(self, draft_id: str, attachment_id: str) -> None:
        try:
            self.client.write(
                "DELETE",
                f"/messages/{segment(draft_id)}/attachments/{segment(attachment_id)}",
                retry="none",
                purpose="Outlook attachment cleanup",
            )
        except M365Error as exc:
            if exc.code == "item_not_found" or exc.ambiguous:
                return  # gone, or the next listing shows whether it is
            raise

    def abandon(self, draft_id: str) -> M365Error:
        """Taint the draft so no run ever sends it, then try to delete it."""
        try:
            self.client.write(
                "PATCH",
                f"/messages/{segment(draft_id)}",
                json_body={
                    "singleValueExtendedProperties": [{"id": STATE_PROPERTY_ID, "value": ABANDONED}]
                },
                retry="none",
                purpose="Outlook draft abandon",
            )
        except M365Error:
            pass
        deleted = self.delete_message(draft_id)
        where = "and deleted" if deleted else "but could not be deleted; it stays in Drafts"
        return M365Error(
            "attachment_state_unknown",
            f"An attachment upload could not be confirmed, so draft {draft_id} was abandoned {where}. "
            "Nothing was sent; a step retry builds a fresh draft.",
            retryable=True,
            failure_class="provider",
        )

    # step 3: send ----------------------------------------------------------------------

    def send(self, draft_id: str) -> dict[str, Any]:
        report_progress("Sending the email")
        try:
            self.client.write(
                "POST",
                f"/messages/{segment(draft_id)}/send",
                retry="none",
                purpose="Outlook send",
                ok_statuses=(200, 202, 204),
            )
        except M365Error as exc:
            if not exc.ambiguous:
                # A confirmed refusal: nothing went out, so the draft is only clutter.
                self.delete_message(draft_id)
                raise
            confirmation = self.confirm_sent(draft_id)
            if confirmation is None:
                raise M365Error(
                    "timeout_unknown",
                    f"Outlook did not confirm the send of draft {draft_id}; check Sent Items before sending again",
                    http_status=exc.http_status,
                    ambiguous=True,
                ) from exc
            return self._success(*confirmation)
        return self._success(draft_id, "send_accepted")

    def confirm_sent(self, draft_id: str) -> tuple[str, str] | None:
        remaining = getattr(self.client.transport, "remaining_seconds", None)
        budget = CONFIRM_BUDGET_SECONDS
        if callable(remaining):
            budget = min(budget, max(0.0, float(remaining()) - 10.0))
        deadline = clock() + budget
        while True:
            report_progress("Confirming the send")
            try:
                row = self.client.get(
                    f"/messages/{segment(draft_id)}",
                    query={"$select": "id,isDraft,parentFolderId"},
                    response_limit=64 * 1024,
                    purpose="Outlook send confirmation",
                )
                if row.get("isDraft") is False:
                    return draft_id, "immutable_id"
            except M365Error:
                pass
            try:
                sent = self.find_sent()
                if sent:
                    return sent[0], "sent_items_search"
            except M365Error:
                pass
            if clock() + CONFIRM_INTERVAL_SECONDS > deadline:
                return None
            sleep(CONFIRM_INTERVAL_SECONDS)

    # lookups ---------------------------------------------------------------------------

    def _key_filter(self) -> str:
        return (
            f"singleValueExtendedProperties/Any(ep: ep/id eq {odata_string(KEY_PROPERTY_ID)} "
            f"and ep/value eq {odata_string(self.key)})"
        )

    def find_drafts(self) -> list[FoundDraft]:
        rows = self.client.list_all(
            "/mailFolders/drafts/messages",
            query={
                "$filter": self._key_filter(),
                "$select": "id,createdDateTime,isDraft",
                "$expand": f"singleValueExtendedProperties($filter=id eq {odata_string(STATE_PROPERTY_ID)})",
                "$top": "50",
            },
            max_items=200,
            purpose="Outlook draft lookup",
        )
        drafts = []
        for row in rows:
            state = ""
            for prop in row.get("singleValueExtendedProperties") or []:
                if (
                    isinstance(prop, Mapping)
                    and str(prop.get("id") or "").lower() == STATE_PROPERTY_ID.lower()
                ):
                    state = str(prop.get("value") or "")
            drafts.append(
                FoundDraft(provider_id(row.get("id")), str(row.get("createdDateTime") or ""), state)
            )
        return drafts

    def find_sent(self) -> list[str]:
        page = self.client.get(
            "/mailFolders/sentitems/messages",
            query={"$filter": self._key_filter(), "$select": "id,isDraft", "$top": "5"},
            purpose="Outlook Sent Items lookup",
        )
        rows = page.get("value")
        if not isinstance(rows, list):
            raise M365Error(
                "provider_unavailable", "Microsoft Graph returned an invalid message list"
            )
        return [
            provider_id(row.get("id"))
            for row in rows
            if isinstance(row, Mapping) and row.get("isDraft") is not True
        ]

    # cleanup ---------------------------------------------------------------------------

    def delete_message(self, message_id: str) -> bool:
        try:
            self.client.write(
                "DELETE",
                f"/messages/{segment(message_id)}",
                retry="none",
                purpose="Outlook draft cleanup",
            )
        except M365Error as exc:
            return exc.code == "item_not_found"
        return True

    def cleanup_quietly(self) -> None:
        """Steps 1–2 failed for good: delete every draft carrying the key."""
        try:
            for draft in self.find_drafts():
                self.delete_message(draft.message_id)
        except M365Error:
            pass

    def _success(self, message_id: str, confirmation: str) -> dict[str, Any]:
        return {
            "message_id": message_id,
            "confirmation": confirmation,
            "attachment_count": self.attachment_count,
            "external_effect_status": "succeeded",
            "definitely_no_external_effect": False,
        }


# -- helpers -------------------------------------------------------------------------


def _upload_failed(exc: M365Error) -> M365Error:
    if exc.code in {"authorization_required", "rate_limited", "egress_host_not_allowed"}:
        return without_ambiguity(exc)
    return M365Error(
        "attachment_upload_failed",
        f"Outlook did not accept the large attachment ({exc.code}). From a shared mailbox, "
        "see the README for the known Microsoft issue and its workaround.",
        http_status=exc.http_status,
    )


class _Hashing:
    """Pass chunks through while hashing them."""

    def __init__(self, chunks: Iterator[bytes]) -> None:
        self._chunks = chunks
        self._digest = hashlib.sha256()

    def __iter__(self) -> Iterator[bytes]:
        for chunk in self._chunks:
            self._digest.update(chunk)
            yield chunk

    def hexdigest(self) -> str:
        return self._digest.hexdigest()


def _rechunk(chunks: _Hashing, size: int) -> Iterator[tuple[int, bytes]]:
    """Fixed-size byte ranges; the last one is yielded after the source is exhausted."""
    buffer = bytearray()
    offset = 0
    for chunk in chunks:
        buffer.extend(chunk)
        while len(buffer) > size:
            piece = bytes(buffer[:size])
            del buffer[:size]
            yield offset, piece
            offset += len(piece)
    if buffer:
        yield offset, bytes(buffer)


__all__ = [
    "ABANDONED",
    "FORWARD_FIELDS",
    "HEADER_NAME",
    "KEY_PROPERTY_ID",
    "PROPERTY_SET_GUID",
    "REPLY_FIELDS",
    "SEND_FIELDS",
    "STATE_PROPERTY_ID",
    "SendFlow",
    "forward_message",
    "plan_forward",
    "plan_reply",
    "plan_send",
    "reply_message",
    "send_message",
]
