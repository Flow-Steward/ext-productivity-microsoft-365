"""Excel failures that must carry a bounded ``result`` (per-row statuses)."""

from __future__ import annotations

from typing import Any

from m365.errors import M365Error


class ExcelRowsError(M365Error):
    """An :class:`M365Error` whose failure envelope keeps a bounded result.

    ``append_table_rows`` reports ``dedupe_conflict`` with the conflicting rows and
    ``timeout_unknown`` with a per-row ``confirmed``/``unknown`` status. The shared
    framework builds the failure envelope from :meth:`retry_facts`, so the result
    travels there.
    """

    def __init__(self, code: str, message: str, *, result: dict[str, Any], **kwargs: Any) -> None:
        super().__init__(code, message, **kwargs)
        self.result = dict(result)

    def retry_facts(self) -> dict[str, Any]:
        facts = super().retry_facts()
        facts["result"] = dict(self.result)
        return facts


def without_effect(exc: M365Error, purpose: str) -> M365Error:
    """Re-classify a failure of a call that has no external data effect.

    The transport treats every non-GET timeout or 5xx as an ambiguous write. A
    workbook ``createSession`` or a Search API query changes no user data, so its
    failure is a plain, retryable provider or network failure.
    """
    if not exc.ambiguous:
        return exc
    code = "network_failure" if exc.http_status is None else "provider_unavailable"
    return M365Error(
        code,
        f"{purpose} did not complete",
        http_status=exc.http_status,
        retry_after_seconds=exc.retry_after_seconds,
    )


def remap_not_found(exc: M365Error, code: str, message: str) -> M365Error:
    if exc.code == "item_not_found":
        return M365Error(code, message, http_status=exc.http_status)
    return exc


__all__ = ["ExcelRowsError", "remap_not_found", "without_effect"]
