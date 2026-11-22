"""Excel writes: sessions, replace semantics, dedupe and ambiguous outcomes (B5.4, B6)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_excel_support import (  # noqa: E402
    SESSION_ID,
    Fault,
    FakeExcel,
    ambiguous,
    ids_workbook,
    run,
)

from m365.errors import M365Error  # noqa: E402


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    from excel import writes

    waits: list[float] = []
    monkeypatch.setattr(writes, "sleep", waits.append)
    return waits


def _write(fake: FakeExcel, **overrides):
    action_input = {
        "workbook": ids_workbook(),
        "worksheet": "Data",
        "address": "B2",
        "inline_rows": [["Invoice", "Amount"], ["INV-1", 100]],
    }
    action_input.update(overrides)
    return run(fake, "write_range", action_input)


def _assert_no_write_retries(fake: FakeExcel) -> None:
    """B6: every write of these operations bypasses the transport retry path."""
    for call in fake.calls:
        if call["method"] != "GET":
            assert call["retry"] == "none", call


# -- write_range ------------------------------------------------------------


def test_write_range_runs_in_a_persistent_session_and_closes_it() -> None:
    fake = FakeExcel()

    response = _write(fake)

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    methods = [(c["method"], c["path"].rsplit("/", 1)[1]) for c in fake.calls]
    assert methods == [
        ("GET", "01ACMEBOOK"),
        ("POST", "createSession"),
        ("PATCH", "range(address='B2:C3')"),
        ("POST", "closeSession"),
    ]
    patch = fake.calls[2]
    assert patch["headers"] == {"workbook-session-id": SESSION_ID}
    assert patch["body"] == {"values": [["Invoice", "Amount"], ["INV-1", 100]]}
    assert fake.calls[3]["headers"] == {"workbook-session-id": SESSION_ID}
    assert fake.sheets["Data"].cells[(3, 3)] == 100
    assert response["result"] == {
        "workbook": response["result"]["workbook"],
        "worksheet": "Data",
        "address": "B2:C3",
        "row_count": 2,
        "column_count": 2,
        "confirmation": "response",
    }
    _assert_no_write_retries(fake)


def test_write_range_replaces_a_whole_rectangle_and_clears_the_rest() -> None:
    fake = FakeExcel()
    for column in range(1, 4):
        for row in range(1, 4):
            fake.sheets["Data"].put_raw(column, row, "old")

    response = _write(fake, address="A1:C3", inline_rows=[["new", None], ["x"]])

    assert response["ok"] is True, response
    assert fake.calls_matching("PATCH", "range")[0]["body"] == {
        "values": [["new", "", ""], ["x", "", ""], ["", "", ""]]
    }
    assert fake.sheets["Data"].cells == {(1, 1): "new", (1, 2): "x"}


def test_write_range_rejects_values_that_do_not_fit_the_rectangle() -> None:
    fake = FakeExcel()

    response = _write(fake, address="A1:A2", inline_rows=[["a", "b"]])

    assert response["error_code"] == "range_invalid"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


def test_write_range_stores_formula_text_as_text_unless_formulas_are_allowed() -> None:
    fake = FakeExcel()

    _write(fake, inline_rows=[['=WEBSERVICE("https://x")', "+1"]])
    _write(fake, address="A10", inline_rows=[["=1+1"]], allow_formulas=True)

    patches = fake.calls_matching("PATCH", "range")
    assert patches[0]["body"]["values"] == [['\'=WEBSERVICE("https://x")', "+1"]]
    assert patches[1]["body"]["values"] == [["=1+1"]]


def test_ambiguous_write_range_with_equal_values_is_confirmed_by_reread() -> None:
    fake = FakeExcel()
    fake.faults.append(Fault("PATCH", "range(", ambiguous(504), apply=True))

    response = _write(fake, inline_rows=[["INV-1", "100", None], ["=x", True, 2.5]])

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert response["external_effect_status"] == "succeeded"
    assert len(fake.calls_matching("PATCH", "range(")) == 1
    reread = fake.calls_matching("GET", "range(")[0]
    assert reread["query"] == {"$select": "values,formulas"}
    assert reread["headers"] == {"workbook-session-id": SESSION_ID}
    assert fake.calls[-1]["path"].endswith("/closeSession")
    _assert_no_write_retries(fake)


def test_ambiguous_write_range_with_different_values_is_timeout_unknown(no_sleep) -> None:
    fake = FakeExcel()
    fake.sheets["Data"].put_raw(2, 2, "edited by someone")
    fake.faults.append(Fault("PATCH", "range(", ambiguous(503), apply=False))

    response = _write(fake)

    assert response["ok"] is False
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False
    # The write is never repeated; the range is re-read a bounded number of times.
    assert len(fake.calls_matching("PATCH", "range(")) == 1
    assert 1 <= len(fake.calls_matching("GET", "range(")) <= 4
    assert sum(no_sleep) <= 20
    assert fake.calls[-1]["path"].endswith("/closeSession")


def test_ambiguous_write_range_that_lands_late_is_confirmed_on_a_later_reread() -> None:
    fake = FakeExcel()
    fake.faults.append(Fault("PATCH", "range(", ambiguous(), apply=False, land_after=1))

    response = _write(fake)

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert len(fake.calls_matching("GET", "range(")) == 2


def test_reread_is_bounded_by_the_operation_deadline(no_sleep) -> None:
    fake = FakeExcel()
    fake.remaining = 16.0
    fake.faults.append(Fault("PATCH", "range(", ambiguous(), apply=False))

    response = _write(fake)

    assert response["error_code"] == "timeout_unknown"
    assert len(fake.calls_matching("GET", "range(")) == 1
    assert no_sleep == []


def test_definite_write_failure_reports_no_external_effect_and_closes_the_session() -> None:
    fake = FakeExcel()

    response = _write(fake, worksheet="Missing")

    assert response["error_code"] == "worksheet_not_found"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls[-1]["path"].endswith("/closeSession")


def test_session_create_timeout_is_not_an_ambiguous_write() -> None:
    fake = FakeExcel()
    fake.faults.append(Fault("POST", "createSession", ambiguous(None), apply=False))

    response = _write(fake)

    assert response["error_code"] == "network_failure"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls_matching("PATCH", "range(") == []


def test_close_session_failure_does_not_change_a_successful_write() -> None:
    fake = FakeExcel()
    fake.faults.append(Fault("POST", "closeSession", ambiguous(), apply=False))

    assert _write(fake)["ok"] is True


def test_write_range_from_a_dataset_with_header(monkeypatch) -> None:
    from excel import writes

    monkeypatch.setattr(
        writes,
        "read_input_rows",
        lambda _payload, handle, max_rows: (
            ["Invoice", "Amount"],
            [{"Invoice": "INV-1", "Amount": 1}, {"Invoice": "INV-2"}],
        ),
    )
    fake = FakeExcel()

    response = run(
        fake,
        "write_range",
        {
            "workbook": ids_workbook(),
            "worksheet": "Data",
            "address": "A1",
            "dataset_handle": "dataset:acme-in",
            "include_header": True,
        },
    )

    assert response["ok"] is True, response
    assert fake.calls_matching("PATCH", "range(")[0]["body"]["values"] == [
        ["Invoice", "Amount"],
        ["INV-1", 1],
        ["INV-2", ""],
    ]


@pytest.mark.parametrize(
    "overrides",
    [
        {"inline_rows": None},
        {"dataset_handle": "dataset:x"},
        {"include_header": True},
        {"allow_formulas": "yes"},
        {"inline_rows": [[{"nested": 1}]]},
        {"inline_rows": []},
        {"address": None},
        {"worksheet": None},
    ],
)
def test_write_range_input_contract(overrides) -> None:
    fake = FakeExcel()

    response = _write(fake, **overrides)

    assert response["ok"] is False
    assert response["error_code"] in {"invalid_payload", "range_invalid"}
    assert fake.calls == []


@pytest.mark.parametrize("operation", ["write_range", "append_table_rows"])
def test_test_run_suppresses_writes_before_any_graph_call(operation) -> None:
    fake = FakeExcel()
    action_input = {
        "workbook": ids_workbook(),
        "inline_rows": [{"Invoice": "INV-1"}],
        "table": "Invoices",
    }
    if operation == "write_range":
        action_input = {
            "workbook": ids_workbook(),
            "worksheet": "Data",
            "address": "A1",
            "inline_rows": [["x"]],
        }

    response = run(fake, operation, action_input, test_mode=True)

    assert response["ok"] is True
    assert response["external_effect_status"] == "suppressed"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


# -- append_table_rows ------------------------------------------------------


def _append(fake: FakeExcel, rows, **overrides):
    action_input = {"workbook": ids_workbook(), "table": "Invoices", "inline_rows": rows}
    action_input.update(overrides)
    return run(fake, "append_table_rows", action_input)


def _invoices(fake: FakeExcel):
    return fake.add_table("Invoices", ["Invoice", "Customer", "Amount"])


def test_append_uses_one_rows_add_in_header_order_inside_a_session() -> None:
    fake = FakeExcel()
    table = _invoices(fake)
    table.add_raw(["INV-0", "Acme", 5])

    response = _append(
        fake,
        [
            {"Amount": 100, "Invoice": "INV-1"},
            {"Customer": "Acme", "Invoice": "INV-2", "Amount": "=1"},
        ],
    )

    assert response["ok"] is True, response
    adds = fake.calls_matching("POST", "/rows/add")
    assert len(adds) == 1
    assert adds[0]["retry"] == "none"
    assert adds[0]["headers"] == {"workbook-session-id": SESSION_ID}
    assert adds[0]["body"] == {
        "index": None,
        "values": [["INV-1", "", 100], ["INV-2", "Acme", "'=1"]],
    }
    header = fake.calls_matching("GET", "/headerRowRange")[0]
    assert header["headers"] == {"workbook-session-id": SESSION_ID}
    result = response["result"]
    assert result["appended_count"] == 2
    assert result["skipped_count"] == 0
    assert result["changed"] is True
    assert result["confirmation"] == "response"
    assert result["first_row_index"] == 1
    assert [row["status"] for row in result["rows"]] == ["appended", "appended"]
    assert fake.calls[-1]["path"].endswith("/closeSession")
    # Without dedupe_column the table body is not read.
    assert fake.calls_matching("GET", "/dataBodyRange") == []
    _assert_no_write_retries(fake)


def test_append_rejects_unknown_columns_and_more_than_500_rows() -> None:
    fake = FakeExcel()
    _invoices(fake)

    unknown = _append(fake, [{"Invoice": "INV-1", "Total": 3}])
    too_many = _append(fake, [{"Invoice": f"INV-{n}"} for n in range(501)])

    assert unknown["error_code"] == "invalid_payload"
    assert "Total" in unknown["error"]
    assert too_many["error_code"] == "invalid_payload"
    assert fake.calls_matching("POST", "/rows/add") == []


def test_append_dataset_over_500_rows_is_refused(monkeypatch) -> None:
    from excel import writes

    def rows(_payload, _handle, max_rows):
        assert max_rows == 500
        raise M365Error("range_too_large", "too many rows")

    monkeypatch.setattr(writes, "read_input_rows", rows)
    fake = FakeExcel()

    response = run(
        fake,
        "append_table_rows",
        {"workbook": ids_workbook(), "table": "Invoices", "dataset_handle": "dataset:x"},
    )

    assert response["error_code"] == "range_too_large"
    assert fake.calls == []


def test_append_to_a_missing_table_is_table_not_found() -> None:
    fake = FakeExcel()

    response = _append(fake, [{"Invoice": "INV-1"}])

    assert response["error_code"] == "table_not_found"
    assert response["definitely_no_external_effect"] is True


def test_dedupe_equal_key_is_skipped_and_new_key_is_appended() -> None:
    fake = FakeExcel()
    table = _invoices(fake)
    table.add_raw(["INV-1", "Acme", 100])

    response = _append(
        fake,
        [{"Invoice": "INV-1", "Amount": "100"}, {"Invoice": "INV-2", "Amount": 7}],
        dedupe_column="Invoice",
    )

    assert response["ok"] is True, response
    assert fake.calls_matching("POST", "/rows/add")[0]["body"]["values"] == [["INV-2", "", 7]]
    assert response["result"]["rows"] == [
        {"index": 0, "key": "INV-1", "status": "skipped", "reason": "already_present"},
        {"index": 1, "key": "INV-2", "status": "appended"},
    ]
    assert response["result"]["skipped_count"] == 1
    assert response["result"]["appended_count"] == 1
    body_read = fake.calls_matching("GET", "/dataBodyRange")[0]
    assert body_read["headers"] == {"workbook-session-id": SESSION_ID}


def test_dedupe_all_present_writes_nothing() -> None:
    fake = FakeExcel()
    table = _invoices(fake)
    table.add_raw(["INV-1", "Acme", 100])

    response = _append(fake, [{"Invoice": "INV-1", "Customer": "Acme"}], dedupe_column="Invoice")

    assert response["ok"] is True
    assert response["result"]["changed"] is False
    assert response["result"]["confirmation"] == "none"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls_matching("POST", "/rows/add") == []


def test_dedupe_conflict_fails_lists_rows_and_writes_nothing() -> None:
    fake = FakeExcel()
    table = _invoices(fake)
    table.add_raw(["INV-1", "Acme", 100])
    table.add_raw(["INV-3", "Acme", 9])

    response = _append(
        fake,
        [
            {"Invoice": "INV-1", "Amount": 999},
            {"Invoice": "INV-2", "Amount": 1},
            {"Invoice": "INV-3", "Amount": 9},
        ],
        dedupe_column="Invoice",
    )

    assert response["ok"] is False
    assert response["error_code"] == "dedupe_conflict"
    assert "INV-1" in response["error"]
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert response["retryable"] is False
    assert response["result"]["rows"] == [
        {"index": 0, "key": "INV-1", "status": "conflict"},
        {"index": 1, "key": "INV-2", "status": "not_written"},
        {"index": 2, "key": "INV-3", "status": "skipped"},
    ]
    assert fake.calls_matching("POST", "/rows/add") == []
    assert table.count == 2


@pytest.mark.parametrize(
    "rows",
    [
        [{"Invoice": "INV-1"}, {"Invoice": "INV-1"}],
        [{"Invoice": "7"}, {"Invoice": 7}],
        [{"Invoice": "INV-1"}, {"Amount": 3}],
    ],
)
def test_dedupe_repeated_or_missing_key_in_one_batch_is_invalid_before_graph(rows) -> None:
    fake = FakeExcel()
    _invoices(fake)

    response = _append(fake, rows, dedupe_column="Invoice")

    assert response["error_code"] == "invalid_payload"
    assert fake.calls == []


def test_dedupe_column_must_be_a_table_column() -> None:
    fake = FakeExcel()
    _invoices(fake)

    response = _append(fake, [{"Invoice": "INV-1"}], dedupe_column="Total")

    assert response["error_code"] == "invalid_payload"
    assert fake.calls_matching("POST", "/rows/add") == []


def test_ambiguous_append_with_all_keys_confirmed_is_success() -> None:
    fake = FakeExcel()
    _invoices(fake)
    fake.faults.append(Fault("POST", "/rows/add", ambiguous(504), apply=True))

    response = _append(
        fake,
        [{"Invoice": "INV-1", "Amount": 1}, {"Invoice": "INV-2", "Amount": 2}],
        dedupe_column="Invoice",
    )

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert [row["status"] for row in response["result"]["rows"]] == ["appended", "appended"]
    assert len(fake.calls_matching("POST", "/rows/add")) == 1
    _assert_no_write_retries(fake)


def test_ambiguous_append_with_partial_confirmation_is_timeout_unknown_per_row(no_sleep) -> None:
    fake = FakeExcel()
    table = _invoices(fake)
    table.add_raw(["INV-0", "Acme", 5])
    # Only the first new row becomes visible after the timeout; the second never does
    # within the bound.
    fake.faults.append(
        Fault(
            "POST",
            "/rows/add",
            ambiguous(503),
            apply=False,
            on_raise=lambda: fake.delayed.append((1, lambda: table.add(["INV-1", "", 1]))),
        )
    )

    response = _append(
        fake,
        [
            {"Invoice": "INV-0", "Customer": "Acme", "Amount": 5},
            {"Invoice": "INV-1", "Amount": 1},
            {"Invoice": "INV-2", "Amount": 2},
        ],
        dedupe_column="Invoice",
    )

    assert response["ok"] is False
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False
    assert response["result"]["rows"] == [
        {"index": 0, "key": "INV-0", "status": "skipped", "reason": "already_present"},
        {"index": 1, "key": "INV-1", "status": "confirmed"},
        {"index": 2, "key": "INV-2", "status": "unknown"},
    ]
    assert len(fake.calls_matching("POST", "/rows/add")) == 1
    assert 1 < len(fake.calls_matching("GET", "/dataBodyRange")) <= 5
    assert no_sleep
    assert fake.calls[-1]["path"].endswith("/closeSession")


def test_ambiguous_append_without_dedupe_is_always_timeout_unknown() -> None:
    fake = FakeExcel()
    _invoices(fake)
    fake.faults.append(Fault("POST", "/rows/add", ambiguous(504), apply=True))

    response = _append(fake, [{"Invoice": "INV-1"}])

    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["result"]["rows"] == [{"index": 0, "key": "", "status": "unknown"}]
    assert fake.calls_matching("GET", "/dataBodyRange") == []
    assert len(fake.calls_matching("POST", "/rows/add")) == 1


# -- B6: no transport retries for Excel writes, with the real transport -------


class _Response:
    def __init__(self, status: int, body: bytes) -> None:
        self.status = status
        self.headers: dict[str, str] = {}
        self._body = body

    def read(self, size: int = -1) -> bytes:
        chunk, self._body = (
            (self._body, b"") if size < 0 else (self._body[:size], self._body[size:])
        )
        return chunk

    def getcode(self) -> int:
        return self.status

    def __enter__(self):
        return self

    def __exit__(self, *_args) -> None:
        return None


def _opener(routes, seen):
    import io
    import json
    from urllib.error import HTTPError

    def opener(request, **_kwargs):
        method = request.get_method()
        url = request.full_url
        seen.append((method, url))
        for (route_method, marker), outcome in routes.items():
            if method == route_method and marker in url:
                status, body = outcome(len([s for s in seen if s == (method, url)]))
                if status >= 400:
                    raise HTTPError(url, status, "error", {"Retry-After": "0"}, io.BytesIO(b"{}"))
                return _Response(status, json.dumps(body).encode())
        raise AssertionError(f"unexpected {method} {url}")

    return opener


def _graph_routes(write_marker: tuple[str, str], *, read_values) -> dict:
    from test_excel_support import DRIVE_ID, ITEM_ID

    item = {
        "id": ITEM_ID,
        "name": "Acme.xlsx",
        "file": {},
        "parentReference": {"driveId": DRIVE_ID, "driveType": "documentLibrary"},
    }
    return {
        ("GET", f"/items/{ITEM_ID}?"): lambda _n: (200, item),
        ("POST", "/createSession"): lambda _n: (201, {"id": SESSION_ID}),
        ("POST", "/closeSession"): lambda _n: (204, {}),
        write_marker: lambda _n: (503, {}),
        ("GET", "/headerRowRange"): lambda _n: (200, {"values": [["Invoice"]]}),
        ("GET", "/dataBodyRange"): lambda _n: (200, {"address": "Data!A2:A2"}),
        ("GET", "/range(address="): lambda _n: (200, read_values),
    }


@pytest.mark.parametrize(
    ("operation", "action_input", "write_marker", "read_values"),
    [
        (
            "write_range",
            {"worksheet": "Data", "address": "A1", "inline_rows": [["INV-1"]]},
            ("PATCH", "/range(address="),
            {"values": [["other"]], "formulas": [["other"]]},
        ),
        (
            "append_table_rows",
            {
                "table": "Invoices",
                "inline_rows": [{"Invoice": "INV-1"}],
                "dedupe_column": "Invoice",
            },
            ("POST", "/rows/add"),
            {"values": [[""]], "formulas": [[""]]},
        ),
    ],
)
def test_excel_writes_never_reach_the_transport_retry_path(
    operation, action_input, write_marker, read_values
) -> None:
    from excel.operations import OPERATIONS
    from m365.runtime import run_operation
    from m365.transport import GraphTransport
    from test_excel_support import payload

    seen: list[tuple[str, str]] = []
    opener = _opener(_graph_routes(write_marker, read_values=read_values), seen)

    def factory(_context):
        return GraphTransport(
            "acme-token", opener=opener, sleep=lambda _s: None, deadline_seconds=200
        )

    response = run_operation(
        payload(operation, {"workbook": ids_workbook(), **action_input}),
        OPERATIONS,
        transport_factory=factory,
    )

    assert response["error_code"] == "timeout_unknown", response
    assert response["external_effect_status"] == "timeout_unknown"
    writes = [s for s in seen if s[0] == write_marker[0] and write_marker[1] in s[1]]
    assert len(writes) == 1
    assert len([s for s in seen if s[1].endswith("/createSession")]) == 1
