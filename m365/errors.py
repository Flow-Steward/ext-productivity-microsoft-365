"""Closed error model shared by every Microsoft 365 service package."""

from __future__ import annotations

from typing import Any

SAFE_ERROR_CODES = frozenset(
    {
        "invalid_payload",
        "invalid_connection",
        "invalid_cursor",
        "authorization_required",
        "permission_denied",
        "mailbox_not_found",
        "mailbox_access_denied",
        "invalid_client",
        "client_secret_expired",
        "tenant_not_multitenant",
        "message_not_found",
        "attachment_not_found",
        "attachment_type_unsupported",
        "attachment_too_large",
        "attachment_upload_failed",
        "attachment_state_unknown",
        "folder_not_found",
        "item_not_found",
        "drive_not_found",
        "site_not_found",
        "name_conflict",
        "file_too_large",
        "workbook_not_found",
        "workbook_unsupported",
        "workbook_link_unsupported",
        "worksheet_not_found",
        "table_not_found",
        "range_invalid",
        "dedupe_conflict",
        "search_unsupported",
        "sync_reset",
        "rate_limited",
        "provider_unavailable",
        "network_failure",
        "egress_host_not_allowed",
        "artifact_input_unavailable",
        "artifact_output_unavailable",
        "response_too_large",
        "timeout_unknown",
        "internal_error",
    }
)

#: Codes that describe a provider-side or transient condition the workflow
#: retry policy may retry when nothing external happened.
_RETRYABLE_CLASS = {
    "rate_limited": "provider",
    "provider_unavailable": "provider",
    "network_failure": "transient",
}


class M365Error(Exception):
    """An error with a closed code and the facts Core needs for retries."""

    def __init__(
        self,
        code: str,
        message: str,
        *,
        http_status: int | None = None,
        retry_after_seconds: float | None = None,
        ambiguous: bool = False,
        retryable: bool | None = None,
        failure_class: str | None = None,
    ) -> None:
        self.code = code if code in SAFE_ERROR_CODES else "internal_error"
        self.message = str(message or "The Microsoft 365 operation failed")
        self.http_status = http_status
        self.retry_after_seconds = retry_after_seconds
        #: The request may have taken effect: a write must report timeout_unknown.
        self.ambiguous = ambiguous
        default_class = _RETRYABLE_CLASS.get(self.code)
        self.failure_class = failure_class or default_class
        self.retryable = retryable if retryable is not None else default_class is not None
        super().__init__(f"{self.code}: {self.message}")

    def retry_facts(self) -> dict[str, Any]:
        facts: dict[str, Any] = {}
        if self.failure_class:
            facts["failure_class"] = self.failure_class
            facts["provider_error_code"] = self.code
        facts["retryable"] = bool(self.retryable) and not self.ambiguous
        if self.retry_after_seconds is not None:
            facts["retry_after_seconds"] = self.retry_after_seconds
        if self.http_status is not None:
            facts["http_status"] = self.http_status
        return facts


def error_response(code: str, message: str, **facts: Any) -> dict[str, Any]:
    safe = code if code in SAFE_ERROR_CODES else "internal_error"
    response: dict[str, Any] = {
        "ok": False,
        "result": {},
        "error_code": safe,
        "error": message,
        "errors": [{"code": safe, "message": message}],
    }
    response.update(facts)
    return response


def error_from_exception(exc: M365Error, *, effectful: bool) -> dict[str, Any]:
    """Build the failure envelope; writes also state their external effect."""
    response = error_response(exc.code, exc.message, **exc.retry_facts())
    if effectful:
        if exc.ambiguous:
            response["external_effect_status"] = "timeout_unknown"
            response["definitely_no_external_effect"] = False
        else:
            response["external_effect_status"] = "failed"
            response["definitely_no_external_effect"] = True
    return response


__all__ = ["SAFE_ERROR_CODES", "M365Error", "error_from_exception", "error_response"]
