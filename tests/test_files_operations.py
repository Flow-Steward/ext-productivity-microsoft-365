"""Files operations: validation, drive resolution, listing, search, folders, links, queries."""

from __future__ import annotations

import pytest
from test_files_support import (
    APP_CONNECTION,
    DRIVE_ID,
    GRAPH,
    SITE_ID,
    SITE_URL,
    TOKEN,
    USER_CONNECTION,
    USER_SITE_CONNECTION,
    Harness,
    drive_item,
    graph_error,
    upload_artifacts,
)

SITE = {"id": SITE_ID, "displayName": "Finance", "webUrl": SITE_URL}
SITE_LOOKUP = "/sites/acme.sharepoint.com:/sites/Finance"
APP_BASE = f"/drives/{DRIVE_ID}"


def _ok(response: dict) -> dict:
    assert response["ok"] is True, response
    return response["result"]


def _error(response: dict, code: str) -> dict:
    assert response["ok"] is False, response
    assert response["error_code"] == code, response
    return response


# -- validation ---------------------------------------------------------------


def test_unknown_input_field_is_refused_before_any_request() -> None:
    harness = Harness()
    _error(harness.run("list_drive_items", {"folder": "/x"}), "invalid_payload")
    assert harness.graph.sent == []


def test_item_id_and_item_path_are_mutually_exclusive() -> None:
    harness = Harness()
    response = harness.run("download_file", {"item_id": "01ACME", "item_path": "/a.pdf"})
    _error(response, "invalid_payload")
    _error(harness.run("download_file", {}), "invalid_payload")
    assert harness.graph.sent == []


@pytest.mark.parametrize(
    "filename", ["a/b.pdf", "a\\b.pdf", "x:y.pdf", "..", "report..pdf", " lead.pdf", "q?.pdf", ""]
)
def test_unsafe_upload_names_are_refused_before_artifact_or_provider(filename: str) -> None:
    harness = Harness()
    response = harness.run(
        "upload_file",
        {"artifact_handle": "artifact:cmr-scan", "filename": filename},
        artifacts=upload_artifacts(4),
    )
    _error(response, "invalid_payload")
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert harness.graph.sent == []


def test_invalid_conflict_behavior_is_refused() -> None:
    harness = Harness()
    response = harness.run(
        "upload_file",
        {"artifact_handle": "artifact:cmr-scan", "filename": "a.pdf", "conflict_behavior": "keep"},
        artifacts=upload_artifacts(4),
    )
    _error(response, "invalid_payload")


def test_test_run_suppresses_every_write_before_any_request() -> None:
    harness = Harness()
    for action_id, operation_input in (
        ("upload_file", {"artifact_handle": "artifact:cmr-scan", "filename": "cmr.pdf"}),
        ("create_folder", {"name": "Processed"}),
        ("create_share_link", {"item_id": "01ACME"}),
    ):
        response = harness.run(
            action_id, operation_input, artifacts=upload_artifacts(4), test_mode=True
        )
        assert response["ok"] is True
        assert response["external_effect_status"] == "suppressed"
        assert response["definitely_no_external_effect"] is True
    assert harness.graph.sent == []
    assert harness.transports == []


# -- drive resolution ---------------------------------------------------------


def test_delegated_default_is_the_users_onedrive() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        "/me/drive/root/children",
        (
            200,
            {
                "value": [
                    drive_item("01FOLDER", "Invoices", folder=True, parent_path="/drive/root:"),
                    drive_item("01FILE", "cmr 1.pdf", size=7, parent_path="/drive/root:/Invoices"),
                ]
            },
        ),
    )
    result = _ok(harness.run("list_drive_items", {}))
    assert [row["item_id"] for row in result["items"]] == ["01FOLDER", "01FILE"]
    assert result["items"][0]["path"] == "/Invoices"
    assert result["items"][0]["is_folder"] is True
    assert result["items"][1]["path"] == "/Invoices/cmr 1.pdf"
    assert result["items"][1]["drive_id"] == DRIVE_ID
    assert result["next_cursor"] is None
    sent = harness.graph.sent[0]
    assert sent.headers["authorization"] == f"Bearer {TOKEN}"
    assert "%24top=50" in sent.url
    assert {call[2] for call in harness.calls} == {"read"}


def test_delegated_site_url_resolves_site_then_its_default_library() -> None:
    harness = Harness()
    harness.graph.on("GET", SITE_LOOKUP, (200, SITE))
    harness.graph.on("GET", f"/sites/{SITE_ID}/drive/root/children", (200, {"value": []}))
    _ok(harness.run("list_drive_items", {}, connection=USER_SITE_CONNECTION))
    urls = [sent.url for sent in harness.graph.sent]
    assert urls[0].startswith(f"{GRAPH}{SITE_LOOKUP}?")
    assert urls[1].startswith(f"{GRAPH}/sites/{SITE_ID}/drive/root/children?")


def test_delegated_drive_id_target_and_unknown_site() -> None:
    harness = Harness()
    connection = {**USER_CONNECTION, "config": {"drive_target": "b!otherDrive_2"}}
    harness.graph.on("GET", "/drives/b!otherDrive_2/root/children", (200, {"value": []}))
    _ok(harness.run("list_drive_items", {}, connection=connection))

    missing = Harness()
    missing.graph.on("GET", SITE_LOOKUP, graph_error(404, "itemNotFound"))
    _error(missing.run("list_drive_items", {}, connection=USER_SITE_CONNECTION), "site_not_found")

    bad = Harness()
    connection = {**USER_CONNECTION, "config": {"drive_target": "not a drive"}}
    _error(bad.run("list_drive_items", {}, connection=connection), "invalid_connection")
    assert bad.graph.sent == []


def test_app_only_uses_the_configured_library_and_folder_path() -> None:
    harness = Harness()
    harness.graph.on("GET", f"{APP_BASE}/root:/Invoices/CMR%202026:/children", (200, {"value": []}))
    _ok(
        harness.run(
            "list_drive_items", {"folder_path": "/Invoices/CMR 2026/"}, connection=APP_CONNECTION
        )
    )
    assert len(harness.graph.sent) == 1


def test_app_only_without_a_chosen_library_is_an_invalid_connection() -> None:
    harness = Harness()
    config = {key: value for key, value in APP_CONNECTION["config"].items() if key != "drive_id"}
    response = harness.run("list_drive_items", {}, connection={**APP_CONNECTION, "config": config})
    _error(response, "invalid_connection")
    assert harness.graph.sent == []


def test_missing_folder_is_folder_not_found() -> None:
    harness = Harness()
    harness.graph.on("GET", "/me/drive/items/01GONE/children", graph_error(404, "itemNotFound"))
    _error(harness.run("list_drive_items", {"folder_id": "01GONE"}), "folder_not_found")


# -- test_connection ----------------------------------------------------------


def test_connection_reports_drive_and_quota_for_delegated_onedrive() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        "/me/drive?",
        (
            200,
            {
                "id": DRIVE_ID,
                "name": "OneDrive",
                "driveType": "business",
                "webUrl": "https://acme-my.sharepoint.com/personal/ann/Documents",
                "quota": {"total": 100, "used": 10, "remaining": 90, "state": "normal"},
            },
        ),
    )
    result = _ok(harness.run("test_files_connection", {}))
    assert result["mode"] == "user"
    assert result["drive_id"] == DRIVE_ID
    assert result["quota"] == {"total": 100, "used": 10, "remaining": 90, "state": "normal"}
    assert result["site_id"] is None


@pytest.mark.parametrize(
    ("drive_url", "ok"),
    [
        ("https://acme.sharepoint.com/sites/Finance/Shared%20Documents", True),
        ("https://acme.sharepoint.com/sites/Other/Shared%20Documents", False),
    ],
)
def test_app_connection_test_checks_the_library_belongs_to_the_site(
    drive_url: str, ok: bool
) -> None:
    harness = Harness()
    harness.graph.on("GET", SITE_LOOKUP, (200, SITE))
    harness.graph.on(
        "GET",
        f"{APP_BASE}?",
        (
            200,
            {
                "id": DRIVE_ID,
                "name": "Documents",
                "driveType": "documentLibrary",
                "webUrl": drive_url,
            },
        ),
    )
    response = harness.run("test_files_connection", {}, connection=APP_CONNECTION)
    if ok:
        result = _ok(response)
        assert result["site_id"] == SITE_ID
        assert result["drive_name"] == "Documents"
    else:
        _error(response, "invalid_connection")


# -- cursors ------------------------------------------------------------------


def _paged_harness() -> Harness:
    harness = Harness()
    harness.graph.on(
        "GET",
        "skiptoken=acme2",
        (200, {"value": [drive_item("04", "d.pdf")]}),
    )
    harness.graph.on(
        "GET",
        "/me/drive/root/children",
        (
            200,
            {
                "value": [
                    drive_item("01", "a.pdf"),
                    drive_item("02", "b.pdf"),
                    drive_item("03", "c.pdf"),
                ],
                "@odata.nextLink": f"{GRAPH}/me/drive/root/children?$skiptoken=acme2",
            },
        ),
    )
    return harness


def test_cursor_resumes_inside_a_page_and_follows_next_links() -> None:
    harness = _paged_harness()
    first = _ok(harness.run("list_drive_items", {"limit": 2}))
    assert [row["item_id"] for row in first["items"]] == ["01", "02"]
    cursor = first["next_cursor"]
    assert cursor and "skiptoken" not in cursor and "graph.microsoft.com" not in cursor
    second = _ok(harness.run("list_drive_items", {"limit": 2, "cursor": cursor}))
    assert [row["item_id"] for row in second["items"]] == ["03", "04"]
    assert second["next_cursor"] is None


def test_cursor_is_bound_to_its_input_and_cannot_be_forged() -> None:
    harness = _paged_harness()
    cursor = _ok(harness.run("list_drive_items", {"limit": 2}))["next_cursor"]
    _error(
        harness.run("list_drive_items", {"limit": 3, "cursor": cursor}),
        "invalid_cursor",
    )
    _error(
        harness.run("list_drive_items", {"limit": 2, "cursor": cursor[:-4] + "AAAA"}),
        "invalid_cursor",
    )
    _error(
        harness.run("search_drive_files", {"query": "a", "limit": 2, "cursor": cursor}),
        "invalid_cursor",
    )


def test_next_link_on_a_foreign_host_is_refused() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        "/me/drive/root/children",
        (200, {"value": [drive_item("01", "a.pdf")], "@odata.nextLink": "https://evil.example/x"}),
    )
    _error(harness.run("list_drive_items", {}), "provider_unavailable")


# -- search -------------------------------------------------------------------


def test_app_only_search_filters_names_inside_one_folder_without_recursion() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        f"{APP_BASE}/root:/Inbound:/children",
        (
            200,
            {
                "value": [
                    drive_item("01", "CMR-0042.pdf"),
                    drive_item("02", "cmr-0043.xlsx"),
                    drive_item("03", "invoice.pdf"),
                    drive_item("04", "cmr archive", folder=True),
                ]
            },
        ),
    )
    result = _ok(
        harness.run(
            "search_drive_files",
            {"query": "cmr", "folder_path": "/Inbound", "file_types": [".PDF"]},
            connection=APP_CONNECTION,
        )
    )
    assert [row["item_id"] for row in result["items"]] == ["01"]
    assert result["search_mode"] == "folder_name_filter"
    assert not harness.graph.requests(part="/search(")
    assert "%24top=200" in harness.graph.sent[0].url


def test_delegated_search_uses_drive_search_and_keeps_results_under_the_folder() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        "/me/drive/root/search(q='O%27%27Brien%20cmr')",
        (
            200,
            {
                "value": [
                    drive_item("01", "a.pdf", parent_path="/drive/root:/Invoices/2026"),
                    drive_item("02", "b.pdf", parent_path="/drive/root:/Invoices Old"),
                    drive_item("03", "c.pdf", parent_path="/drive/root:/Other"),
                    drive_item("04", "Invoices", folder=True, parent_path="/drive/root:"),
                ]
            },
        ),
    )
    result = _ok(
        harness.run("search_drive_files", {"query": "O'Brien cmr", "folder_path": "/invoices"})
    )
    assert [row["item_id"] for row in result["items"]] == ["01"]
    assert result["search_mode"] == "drive_search"


def test_delegated_search_rejects_dot_dot_query() -> None:
    harness = Harness()
    _error(harness.run("search_drive_files", {"query": "a..b"}), "invalid_payload")
    assert harness.graph.sent == []


# -- create_folder ------------------------------------------------------------


def test_create_folder_creates_with_fail_conflict_and_idempotent_retry() -> None:
    harness = Harness()
    harness.graph.on(
        "POST",
        f"{APP_BASE}/root:/Invoices:/children",
        (
            201,
            drive_item("01NEW", "Processed", folder=True, parent_path=f"{APP_BASE}/root:/Invoices"),
        ),
    )
    response = harness.run(
        "create_folder",
        {"parent_folder_path": "/Invoices", "name": "Processed"},
        connection=APP_CONNECTION,
    )
    result = _ok(response)
    assert result["changed"] is True
    assert result["path"] == "/Invoices/Processed"
    assert response["external_effect_status"] == "succeeded"
    body = harness.graph.sent[0].json()
    assert body == {"name": "Processed", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    assert [call[2] for call in harness.calls if call[0] == "POST"] == ["idempotent"]


def test_create_folder_name_conflict_with_existing_folder_is_unchanged_success() -> None:
    harness = Harness()
    harness.graph.on("POST", "/me/drive/root/children", graph_error(409, "nameAlreadyExists"))
    harness.graph.on(
        "GET", "/me/drive/root:/Processed:", (200, drive_item("01OLD", "Processed", folder=True))
    )
    result = _ok(harness.run("create_folder", {"name": "Processed"}))
    assert result["changed"] is False
    assert result["item_id"] == "01OLD"


def test_create_folder_conflict_with_a_file_fails_without_effect() -> None:
    harness = Harness()
    harness.graph.on("POST", "/me/drive/root/children", graph_error(409, "nameAlreadyExists"))
    harness.graph.on("GET", "/me/drive/root:/Processed:", (200, drive_item("01F", "Processed")))
    response = _error(harness.run("create_folder", {"name": "Processed"}), "name_conflict")
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True


def test_create_folder_in_missing_parent_is_folder_not_found() -> None:
    harness = Harness()
    harness.graph.on("POST", "/me/drive/items/01GONE/children", graph_error(404, "itemNotFound"))
    _error(
        harness.run("create_folder", {"parent_folder_id": "01GONE", "name": "x"}),
        "folder_not_found",
    )


# -- create_share_link --------------------------------------------------------


@pytest.mark.parametrize(("status", "changed"), [(201, True), (200, False)])
def test_create_share_link_never_retries_and_reports_reuse(status: int, changed: bool) -> None:
    harness = Harness()
    harness.graph.on(
        "POST",
        "/me/drive/items/01ACME/createLink",
        (
            status,
            {
                "id": "perm-1",
                "link": {
                    "type": "view",
                    "scope": "organization",
                    "webUrl": "https://acme.sharepoint.com/:b:/s/x",
                },
            },
        ),
    )
    result = _ok(harness.run("create_share_link", {"item_id": "01ACME"}))
    assert result == {
        "permission_id": "perm-1",
        "link_url": "https://acme.sharepoint.com/:b:/s/x",
        "link_type": "view",
        "scope": "organization",
        "changed": changed,
    }
    assert harness.graph.sent[0].json() == {"type": "view", "scope": "organization"}
    assert [call[2] for call in harness.calls] == ["none"]


def test_create_share_link_ambiguous_failure_is_timeout_unknown_without_repeat() -> None:
    harness = Harness()
    harness.graph.on("POST", "/createLink", graph_error(503))
    response = _error(
        harness.run("create_share_link", {"item_path": "/a.pdf", "link_type": "edit"}),
        "timeout_unknown",
    )
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert len(harness.graph.requests("POST")) == 1


# -- list_site_libraries ------------------------------------------------------


def test_list_site_libraries_resolves_site_url_and_lists_every_library() -> None:
    harness = Harness()
    harness.graph.on("GET", SITE_LOOKUP, (200, SITE))
    harness.graph.on(
        "GET",
        f"/sites/{SITE_ID}/drives",
        (
            200,
            {
                "value": [
                    {
                        "id": DRIVE_ID,
                        "name": "Documents",
                        "driveType": "documentLibrary",
                        "webUrl": f"{SITE_URL}/Shared%20Documents",
                    },
                    {
                        "id": "b!acmeCmr_02",
                        "name": "CMR",
                        "driveType": "documentLibrary",
                        "webUrl": f"{SITE_URL}/CMR",
                    },
                ]
            },
        ),
    )
    result = _ok(
        harness.run("list_site_libraries", {"site_url": SITE_URL + "/"}, connection=APP_CONNECTION)
    )
    assert result["site_id"] == SITE_ID
    assert [row["drive_id"] for row in result["libraries"]] == [DRIVE_ID, "b!acmeCmr_02"]
    assert result["total"] == 2


@pytest.mark.parametrize(
    "site_url",
    [
        "http://acme.sharepoint.com/sites/Finance",
        "https://acme.example.com/sites/Finance",
        "https://user@acme.sharepoint.com/sites/Finance",
        "https://acme.sharepoint.com/sites/Finance?x=1",
    ],
)
def test_list_site_libraries_refuses_non_sharepoint_urls(site_url: str) -> None:
    harness = Harness()
    _error(
        harness.run("list_site_libraries", {"site_url": site_url}, connection=APP_CONNECTION),
        "invalid_payload",
    )
    assert harness.graph.sent == []


def test_list_site_libraries_falls_back_to_the_configured_site() -> None:
    harness = Harness()
    harness.graph.on("GET", SITE_LOOKUP, (200, SITE))
    harness.graph.on("GET", f"/sites/{SITE_ID}/drives", (200, {"value": []}))
    result = _ok(harness.run("list_site_libraries", {}, connection=APP_CONNECTION))
    assert result["libraries"] == []


# -- retry allow-list (B6) ----------------------------------------------------


def test_only_reads_and_create_folder_may_use_transport_retries() -> None:
    from files.operations import OPERATIONS

    writes = {op_id for op_id, op in OPERATIONS.items() if op.effectful}
    assert writes == {"upload_file", "create_folder", "create_share_link"}


# -- resource queries ---------------------------------------------------------


def test_query_ids_are_one_per_connection_type() -> None:
    from files.operations import QUERY_IDS, QUERY_TYPES

    assert QUERY_IDS == {
        "drive_folders_user",
        "drive_folders_app",
        "site_libraries_user",
        "site_libraries_app",
        "site_search_user",
    }
    assert QUERY_TYPES["site_search_user"] == "m365_files_user"


def test_folder_query_lists_root_folders_for_each_type() -> None:
    harness = Harness()
    harness.graph.on(
        "GET",
        f"{APP_BASE}/root/children",
        (
            200,
            {
                "value": [
                    drive_item("01", "Invoices", folder=True, parent_path=f"{APP_BASE}/root:"),
                    drive_item("02", "a.pdf"),
                ]
            },
        ),
    )
    response = harness.query("drive_folders_app", connection=APP_CONNECTION)
    assert response == {
        "ok": True,
        "result": {
            "items": [{"id": "01", "label": "/Invoices", "path": "/Invoices"}],
            "next_cursor": "",
            "total": 1,
        },
    }


def test_folder_query_lists_subfolders_of_a_typed_path() -> None:
    harness = Harness()
    harness.graph.on("GET", "/me/drive/root:/Invoices:/children", (200, {"value": []}))
    assert harness.query("drive_folders_user", search="/Invoices")["ok"] is True


def test_query_refuses_a_connection_of_the_other_type() -> None:
    harness = Harness()
    response = harness.query("drive_folders_app", connection=USER_CONNECTION)
    assert response["ok"] is False
    assert response["error_code"] == "invalid_connection"
    assert harness.graph.sent == []
    response = harness.query("site_search_user", connection=APP_CONNECTION, search="fin")
    assert response["error_code"] == "invalid_connection"


def test_site_search_query_is_delegated_only_and_returns_site_urls() -> None:
    harness = Harness()
    harness.graph.on("GET", "/sites?search=fin", (200, {"value": [SITE]}))
    response = harness.query("site_search_user", search="fin")
    assert response["result"]["items"] == [{"id": SITE_URL, "label": "Finance", "site_id": SITE_ID}]
    empty = Harness()
    assert empty.query("site_search_user")["result"]["items"] == []
    assert empty.graph.sent == []


def test_library_query_uses_the_configured_site_of_an_app_connection() -> None:
    harness = Harness()
    harness.graph.on("GET", SITE_LOOKUP, (200, SITE))
    harness.graph.on(
        "GET",
        f"/sites/{SITE_ID}/drives",
        (200, {"value": [{"id": DRIVE_ID, "name": "Documents", "webUrl": f"{SITE_URL}/Docs"}]}),
    )
    response = harness.query("site_libraries_app", connection=APP_CONNECTION)
    assert response["result"]["items"] == [
        {"id": DRIVE_ID, "label": "Documents", "web_url": f"{SITE_URL}/Docs"}
    ]


def test_create_folder_is_the_only_write_that_rides_transport_retries() -> None:
    harness = Harness()
    harness.graph.on(
        "POST",
        "/me/drive/root/children",
        graph_error(503),
        (201, drive_item("01N", "Processed", folder=True)),
    )
    assert _ok(harness.run("create_folder", {"name": "Processed"}))["changed"] is True
    assert len(harness.graph.requests("POST")) == 2
