"""``find_workbooks``: Microsoft Search API request and result shape."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_excel_support import FakeExcel, ambiguous, run  # noqa: E402

from m365.errors import M365Error  # noqa: E402

SITE_ID = (
    "acme.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000"
)


def _hit(item_id: str, name: str, drive_id: str = "b!acmeFinance", site_id: str = SITE_ID) -> dict:
    return {
        "hitId": item_id,
        "rank": 1,
        "resource": {
            "@odata.type": "#microsoft.graph.driveItem",
            "id": item_id,
            "name": name,
            "webUrl": f"https://acme.sharepoint.com/sites/finance/Shared%20Documents/{name}",
            "parentReference": {"driveId": drive_id, "siteId": site_id},
        },
    }


def _fake(hits: list[dict], *, more: bool = False, total: int | None = None) -> FakeExcel:
    fake = FakeExcel()
    fake.search_body = {
        "value": [
            {
                "searchTerms": ["ledger"],
                "hitsContainers": [
                    {
                        "hits": hits,
                        "total": len(hits) if total is None else total,
                        "moreResultsAvailable": more,
                    }
                ],
            }
        ]
    }
    fake.drives = {
        "b!acmeFinance": {
            "id": "b!acmeFinance",
            "name": "Documents",
            "driveType": "documentLibrary",
        },
        "b!acmeHome": {"id": "b!acmeHome", "name": "OneDrive", "driveType": "personal"},
    }
    fake.sites = {
        SITE_ID: {
            "id": SITE_ID,
            "displayName": "Finance",
            "webUrl": "https://acme.sharepoint.com/sites/finance",
        }
    }
    return fake


def test_find_workbooks_queries_drive_items_and_returns_ids_site_and_library() -> None:
    fake = _fake(
        [
            _hit("01LEDGER", "Acme Ledger.xlsx"),
            _hit("01NOTES", "Ledger notes.docx"),
            _hit("01MACRO", "Ledger macro.xlsm"),
            _hit("01HOME", "Ledger copy.xlsx", drive_id="b!acmeHome"),
        ]
    )

    response = run(fake, "find_workbooks", {"name": 'Ledger "2026" (draft)'})

    assert response["ok"] is True, response
    search = fake.calls[0]
    assert search["method"] == "POST"
    assert search["path"] == "/search/query"
    request = search["body"]["requests"][0]
    assert request["entityTypes"] == ["driveItem"]
    assert request["query"]["queryString"] == "Ledger 2026 draft filetype:xlsx"
    assert request["from"] == 0
    assert request["size"] == 25
    # A Search API query is a read: transport retries are allowed for it (B6).
    assert search["retry"] == "idempotent"
    assert response["result"]["workbooks"] == [
        {
            "drive_id": "b!acmeFinance",
            "drive_item_id": "01LEDGER",
            "name": "Acme Ledger.xlsx",
            "site": "Finance",
            "site_url": "https://acme.sharepoint.com/sites/finance",
            "library": "Documents",
            "web_url": "https://acme.sharepoint.com/sites/finance/Shared%20Documents/Acme Ledger.xlsx",
            "workbook": {"drive_id": "b!acmeFinance", "drive_item_id": "01LEDGER"},
        }
    ]
    assert response["result"]["next_cursor"] == ""
    # One lookup per distinct drive and site.
    assert len([c for c in fake.calls if c["path"] == "/drives/b!acmeFinance"]) == 1
    assert len([c for c in fake.calls if c["path"].startswith("/sites/")]) == 1


def test_find_workbooks_result_maps_into_the_workbook_input() -> None:
    fake = _fake([_hit("01LEDGER", "Acme Ledger.xlsx")])
    found = run(fake, "find_workbooks", {"name": "ledger"})["result"]["workbooks"][0]

    from excel.workbook import validated_workbook

    assert validated_workbook(found["workbook"]) == found["workbook"]


def test_find_workbooks_pages_with_an_opaque_cursor_bound_to_the_name() -> None:
    fake = _fake([_hit("01LEDGER", "Acme Ledger.xlsx")], more=True, total=40)

    first = run(fake, "find_workbooks", {"name": "ledger", "limit": 1})
    cursor = first["result"]["next_cursor"]
    assert cursor and "from" not in cursor
    assert first["result"]["total"] == 40

    second = run(fake, "find_workbooks", {"name": "ledger", "limit": 1, "cursor": cursor})
    assert second["ok"] is True
    assert fake.calls_matching("POST", "/search/query")[-1]["body"]["requests"][0]["from"] == 1

    other = run(fake, "find_workbooks", {"name": "budget", "cursor": cursor})
    assert other["error_code"] == "invalid_cursor"
    broken = run(fake, "find_workbooks", {"name": "ledger", "cursor": "!!not-a-cursor"})
    assert broken["error_code"] == "invalid_cursor"


def test_find_workbooks_validates_name_and_limit() -> None:
    fake = _fake([])

    assert run(fake, "find_workbooks", {"name": '"()"'})["error_code"] == "invalid_payload"
    assert (
        run(fake, "find_workbooks", {"name": "ledger", "limit": 26})["error_code"]
        == "invalid_payload"
    )
    assert run(fake, "find_workbooks", {})["error_code"] == "invalid_payload"
    assert fake.calls == []


def test_find_workbooks_lookup_failures_leave_names_empty() -> None:
    fake = _fake([_hit("01LEDGER", "Acme Ledger.xlsx", drive_id="b!acmeOther")])

    result = run(fake, "find_workbooks", {"name": "ledger"})["result"]

    assert result["workbooks"][0]["library"] == ""
    assert result["workbooks"][0]["drive_id"] == "b!acmeOther"


def test_find_workbooks_timeout_is_a_plain_retryable_failure() -> None:
    from test_excel_support import Fault

    fake = _fake([])
    fake.faults.append(Fault("POST", "/search/query", ambiguous(None), apply=False))

    response = run(fake, "find_workbooks", {"name": "ledger"})

    assert response["error_code"] == "network_failure"
    assert response["retryable"] is True
    assert "external_effect_status" not in response


def test_find_workbooks_authorization_failure_is_reported() -> None:
    from test_excel_support import Fault

    fake = _fake([])
    fake.faults.append(
        Fault(
            "POST",
            "/search/query",
            M365Error("authorization_required", "again", http_status=401),
            apply=False,
        )
    )

    assert run(fake, "find_workbooks", {"name": "ledger"})["error_code"] == "authorization_required"
