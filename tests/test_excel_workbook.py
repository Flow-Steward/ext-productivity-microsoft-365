"""Excel workbook addressing: the two ``workbook`` forms, links and unsupported files."""

from __future__ import annotations

import base64
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_excel_support import (  # noqa: E402
    DRIVE_ID,
    ITEM_ID,
    SHARING_LINK,
    FakeExcel,
    ids_workbook,
    payload,
    run,
)

from m365.errors import M365Error  # noqa: E402


def test_mapping_form_resolves_ids_and_checks_the_drive_item() -> None:
    fake = FakeExcel()

    response = run(fake, "list_workbook_objects", {"workbook": ids_workbook()})

    assert response["ok"] is True, response
    lookup = fake.calls[0]
    assert lookup["method"] == "GET"
    assert lookup["path"] == f"/drives/{DRIVE_ID}/items/{ITEM_ID}"
    assert lookup["query"] == {"$select": "id,name,file,parentReference,webUrl"}
    assert response["result"]["workbook"] == {
        "drive_id": DRIVE_ID,
        "drive_item_id": ITEM_ID,
        "name": "Acme Ledger.xlsx",
        "web_url": "https://acme.sharepoint.com/sites/finance/Ledger.xlsx",
    }


def test_sharing_link_form_resolves_through_shares_without_padding() -> None:
    fake = FakeExcel()

    response = run(fake, "list_workbook_objects", {"workbook": {"sharing_link": SHARING_LINK}})

    assert response["ok"] is True, response
    encoded = base64.urlsafe_b64encode(SHARING_LINK.encode()).decode().rstrip("=")
    lookup = fake.calls[0]
    assert lookup["path"] == f"/shares/u!{encoded}/driveItem"
    assert "=" not in lookup["path"]
    # The link is only an address: no redeem header is sent.
    assert not any(key.lower() == "prefer" for key in lookup["headers"])
    # The rest of the operation uses the IDs Graph returned for the link.
    assert fake.calls[1]["path"].startswith(f"/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook")
    assert response["result"]["workbook"]["drive_item_id"] == ITEM_ID


@pytest.mark.parametrize(
    "error",
    [
        M365Error("item_not_found", "missing", http_status=404),
        M365Error("invalid_payload", "bad", http_status=400),
    ],
)
def test_rejected_sharing_link_is_workbook_link_unsupported(error: M365Error) -> None:
    fake = FakeExcel()
    fake.share_error = error

    response = run(fake, "list_workbook_objects", {"workbook": {"sharing_link": SHARING_LINK}})

    assert response["ok"] is False
    assert response["error_code"] == "workbook_link_unsupported"
    assert "Copy link" in response["error"]
    assert len(fake.calls) == 1


def test_sharing_link_without_access_is_permission_denied() -> None:
    fake = FakeExcel()
    fake.share_error = M365Error("permission_denied", "denied", http_status=403)

    response = run(fake, "list_workbook_objects", {"workbook": {"sharing_link": SHARING_LINK}})

    assert response["error_code"] == "permission_denied"


@pytest.mark.parametrize(
    "workbook",
    [
        {"drive_id": DRIVE_ID, "drive_item_id": ITEM_ID, "sharing_link": SHARING_LINK},
        {"drive_id": DRIVE_ID, "sharing_link": SHARING_LINK},
        {"drive_id": DRIVE_ID},
        {"drive_item_id": ITEM_ID},
        {},
        {"sharing_link": "http://acme.sharepoint.com/:x:/s/finance/x"},
        {"sharing_link": "not a link"},
        {"drive_id": "../etc", "drive_item_id": ITEM_ID},
        {"drive_id": DRIVE_ID, "drive_item_id": ITEM_ID, "path": "/Ledger.xlsx"},
        "https://acme.sharepoint.com/:x:/s/finance/x",
        None,
    ],
)
def test_mixed_partial_or_malformed_forms_are_invalid_payload(workbook) -> None:
    fake = FakeExcel()
    action_input = {} if workbook is None else {"workbook": workbook}

    response = run(fake, "list_workbook_objects", action_input)

    assert response["ok"] is False
    assert response["error_code"] == "invalid_payload"
    assert fake.calls == []


@pytest.mark.parametrize(
    "item_change",
    [
        {"name": "Acme Ledger.xlsm"},
        {"name": "Acme Ledger.xls"},
        {"name": "Acme Ledger.csv"},
        {"file": None},
        {"parentReference": {"driveId": DRIVE_ID, "driveType": "personal"}},
    ],
)
def test_non_xlsx_folder_or_personal_drive_is_workbook_unsupported(item_change) -> None:
    fake = FakeExcel()
    fake.item.update(item_change)

    response = run(fake, "list_workbook_objects", {"workbook": ids_workbook()})

    assert response["error_code"] == "workbook_unsupported"
    assert len(fake.calls) == 1


def test_unknown_ids_are_workbook_not_found() -> None:
    fake = FakeExcel()

    response = run(
        fake,
        "list_workbook_objects",
        {"workbook": {"drive_id": DRIVE_ID, "drive_item_id": "01GONE"}},
    )

    assert response["error_code"] == "workbook_not_found"


def test_write_with_a_rejected_link_reports_no_external_effect() -> None:
    fake = FakeExcel()
    fake.share_error = M365Error("item_not_found", "missing", http_status=404)

    response = run(
        fake,
        "write_range",
        {
            "workbook": {"sharing_link": SHARING_LINK},
            "worksheet": "Data",
            "address": "A1",
            "inline_rows": [["x"]],
        },
    )

    assert response["error_code"] == "workbook_link_unsupported"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True


def test_find_workbooks_takes_only_a_name() -> None:
    fake = FakeExcel()

    response = run(fake, "find_workbooks", {"name": "ledger", "workbook": ids_workbook()})

    assert response["error_code"] == "invalid_payload"
    assert fake.calls == []


def test_excel_needs_a_signed_in_account() -> None:
    from excel.operations import OPERATIONS
    from m365.runtime import run_operation

    request = payload("list_workbook_objects", {"workbook": ids_workbook()})
    request["action"]["target"]["connection"] = {
        "connection_id": "conn-x",
        "connection_type_id": "m365_files_app",
        "config": {
            "tenant_id": "acme.onmicrosoft.com",
            "client_id": "11111111-2222-3333-4444-555555555555",
            "site_url": "https://acme.sharepoint.com/sites/ops",
            "drive_id": "b!drive",
        },
        "secrets": {"client_secret": "app-secret"},
    }

    response = run_operation(request, OPERATIONS, transport_factory=lambda _c: FakeExcel())

    assert response["error_code"] == "invalid_connection"


def test_list_workbook_objects_lists_worksheets_and_tables() -> None:
    from test_excel_support import Sheet

    fake = FakeExcel()
    fake.sheets["Summary 2026"] = Sheet("Summary 2026")
    fake.add_table("Invoices", ["Invoice", "Amount"])

    response = run(fake, "list_workbook_objects", {"workbook": ids_workbook()})

    result = response["result"]
    assert [sheet["name"] for sheet in result["worksheets"]] == ["Data", "Summary 2026"]
    assert result["tables"] == [
        {
            "id": "{t-Invoices}",
            "name": "Invoices",
            "worksheet": "Data",
            "show_headers": True,
            "show_totals": False,
        }
    ]
    assert result["truncated"] is False
    assert any(call["path"].endswith("/worksheets/Summary%202026/tables") for call in fake.calls)


def test_excel_serves_no_resource_query() -> None:
    from excel.operations import QUERY_IDS, handle_query

    assert QUERY_IDS == frozenset()
    assert handle_query({"mode": "query"})["error_code"] == "invalid_payload"
