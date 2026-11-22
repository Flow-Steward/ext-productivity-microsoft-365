"""A1 ranges, value comparison and ``read_range`` to an NDJSON dataset."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_excel_support import FakeExcel, ids_workbook, run  # noqa: E402

from excel import ranges  # noqa: E402
from m365.errors import M365Error  # noqa: E402

# -- A1 addresses ---------------------------------------------------------


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("A1", (1, 1, 1, 1)),
        ("b2:d20", (2, 2, 4, 20)),
        ("$A$1:$C$3", (1, 1, 3, 3)),
        ("XFD1048576", (16384, 1048576, 16384, 1048576)),
        (" AA10:AB11 ", (27, 10, 28, 11)),
    ],
)
def test_parse_address_accepts_cell_bounded_a1(value, expected) -> None:
    bounds = ranges.parse_address(value)

    assert (bounds.start_column, bounds.start_row, bounds.end_column, bounds.end_row) == expected


@pytest.mark.parametrize(
    "value",
    [
        "",
        "A",
        "1",
        "A:C",
        "1:5",
        "A0",
        "XFE1",
        "A1048577",
        "Sheet1!A1",
        "C3:A1",
        "A1:B2:C3",
        "A1 B2",
        None,
        5,
    ],
)
def test_parse_address_rejects_unbounded_reversed_or_sheet_qualified(value) -> None:
    with pytest.raises(M365Error) as caught:
        ranges.parse_address(value)

    assert caught.value.code == "range_invalid"


def test_bounds_address_and_graph_address_round_trip() -> None:
    sheet, bounds = ranges.parse_graph_address("'Acme ''Q1'' Data'!B2:AA30")

    assert sheet == "Acme 'Q1' Data"
    assert bounds.address() == "B2:AA30"
    assert ranges.range_segment(bounds) == "range(address='B2:AA30')"
    assert ranges.parse_graph_address("Data!C5")[1].address() == "C5"
    with pytest.raises(M365Error):
        ranges.parse_graph_address("C5")


@pytest.mark.parametrize(
    ("expected", "actual", "formula", "equal"),
    [
        ("10", 10, None, True),
        (10, 10.0, None, True),
        ("0.1", 0.1, None, True),
        ("TRUE", True, None, True),
        ("acme", "acme", None, True),
        (None, None, None, True),
        ("", None, None, True),
        ("=SUM(A1:A2)", 3, "=SUM(A1:A2)", True),
        ("=x", "=x", "=x", True),
        ("acme", "Acme", None, False),
        ("2026-11-22", 46348, 46348, False),
        (10, 11, None, False),
        (None, "x", None, False),
    ],
)
def test_cells_equal_compares_after_excel_coercion(expected, actual, formula, equal) -> None:
    assert ranges.cells_equal(expected, actual, formula) is equal


def test_outgoing_cell_stores_formula_like_text_as_text_unless_allowed() -> None:
    cell = ranges.outgoing_cell

    assert cell('=HYPERLINK("x")', allow_formulas=False) == '\'=HYPERLINK("x")'
    assert cell("+49 30 1234", allow_formulas=False) == "'+49 30 1234"
    assert cell("@acme", allow_formulas=False) == "'@acme"
    assert cell("'quoted", allow_formulas=False) == "''quoted"
    assert cell("-5", allow_formulas=False) == "-5"
    assert cell(None, allow_formulas=False) == ""
    assert cell(7, allow_formulas=False) == 7
    assert cell("=1+1", allow_formulas=True) == "=1+1"


# -- read_range -----------------------------------------------------------


@pytest.fixture
def dataset_sink(monkeypatch):
    from excel import datasets, reads

    captured: dict[str, bytes] = {}
    monkeypatch.setattr(reads, "output_grant_limit", lambda _payload: 1024 * 1024)

    def writer(_payload, chunks, **kwargs):
        assert kwargs == {
            "binding_key": "excel_range_dataset",
            "content_type": "application/x-ndjson",
        }
        body = b"".join(chunks)
        captured["body"] = body
        return {"dataset_handle": "dataset:acme-range", "size_bytes": len(body)}

    monkeypatch.setattr(datasets, "write_dataset_stream", writer)
    return captured


def _rows(body: bytes) -> list[dict]:
    return [json.loads(line) for line in body.splitlines()]


def _fill(fake: FakeExcel, rows: list[list], *, start_column: int = 1, start_row: int = 1) -> None:
    sheet = fake.sheets["Data"]
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if value is not None:
                sheet.put_raw(start_column + c, start_row + r, value)


def test_read_range_streams_ndjson_with_headers_schema_and_preview(dataset_sink) -> None:
    fake = FakeExcel()
    _fill(fake, [["SKU", "Qty", None, "SKU"], ["A-1", 10, True, "x"], ["B-2", 2.5, None, None]])

    response = run(
        fake, "read_range", {"workbook": ids_workbook(), "worksheet": "Data", "address": "A1:D3"}
    )

    assert response["ok"] is True, response
    assert _rows(dataset_sink["body"]) == [
        {"SKU": "A-1", "Qty": 10, "column_3": True, "SKU_2": "x"},
        {"SKU": "B-2", "Qty": 2.5, "column_3": None, "SKU_2": None},
    ]
    result = response["result"]
    assert result["dataset_handle"] == "dataset:acme-range"
    assert result["row_count"] == 2
    assert result["column_count"] == 4
    assert result["schema"] == [
        {"name": "SKU", "type": "string", "nullable": True},
        {"name": "Qty", "type": "number", "nullable": True},
        {"name": "column_3", "type": "boolean", "nullable": True},
        {"name": "SKU_2", "type": "string", "nullable": True},
    ]
    assert result["preview"] == _rows(dataset_sink["body"])
    assert result["address"] == "A1:D3"
    reads = fake.calls_matching("GET", "/range(")
    assert [call["path"].rsplit("/", 1)[1] for call in reads] == ["range(address='A1:D3')"]
    assert reads[0]["query"] == {"$select": "values"}
    # Reads use no workbook session.
    assert fake.calls_matching("POST", "createSession") == []


def test_read_range_without_header_and_in_bounded_row_slices(dataset_sink, monkeypatch) -> None:
    from excel import reads

    monkeypatch.setattr(reads, "CELLS_PER_REQUEST", 4)
    fake = FakeExcel()
    _fill(fake, [[1, 2], [3, 4], [5, 6], [7, 8], [9, 10]], start_column=2, start_row=3)

    response = run(
        fake,
        "read_range",
        {
            "workbook": ids_workbook(),
            "worksheet": "Data",
            "address": "B3:C7",
            "header_mode": "none",
        },
    )

    assert response["ok"] is True, response
    assert [c["path"].rsplit("/", 1)[1] for c in fake.calls_matching("GET", "/range(")] == [
        "range(address='B3:C4')",
        "range(address='B5:C6')",
        "range(address='B7:C7')",
    ]
    assert _rows(dataset_sink["body"])[0] == {"column_1": 1, "column_2": 2}
    assert response["result"]["row_count"] == 5


def test_read_range_uses_the_used_range_when_no_address(dataset_sink) -> None:
    fake = FakeExcel()
    _fill(fake, [["Name"], ["Acme"]], start_column=3, start_row=2)

    response = run(fake, "read_range", {"workbook": ids_workbook(), "worksheet": "Data"})

    assert response["result"]["address"] == "C2:C3"
    assert fake.calls_matching("GET", "/usedRange(valuesOnly=true)")
    assert _rows(dataset_sink["body"]) == [{"Name": "Acme"}]


def test_read_range_reads_a_whole_table(dataset_sink) -> None:
    fake = FakeExcel()
    table = fake.add_table("Invoices", ["Invoice", "Amount"])
    table.add_raw(["INV-1", 100])
    table.add_raw(["INV-2", 250])

    response = run(fake, "read_range", {"workbook": ids_workbook(), "table": "Invoices"})

    assert response["ok"] is True, response
    assert _rows(dataset_sink["body"]) == [
        {"Invoice": "INV-1", "Amount": 100},
        {"Invoice": "INV-2", "Amount": 250},
    ]
    assert response["result"]["worksheet"] == "Data"


@pytest.mark.parametrize(
    ("action_input", "code"),
    [
        ({"worksheet": "Data", "table": "Invoices"}, "invalid_payload"),
        ({}, "invalid_payload"),
        ({"table": "Invoices", "address": "A1:B2"}, "invalid_payload"),
        ({"worksheet": "Data", "address": "A:B"}, "range_invalid"),
        ({"worksheet": "Data", "header_mode": "auto"}, "invalid_payload"),
        ({"worksheet": "Missing", "address": "A1:B2"}, "worksheet_not_found"),
        ({"table": "Missing"}, "table_not_found"),
        ({"worksheet": "Data", "address": "A1:Z100000"}, "range_too_large"),
    ],
)
def test_read_range_input_and_lookup_failures(dataset_sink, action_input, code) -> None:
    fake = FakeExcel()

    response = run(fake, "read_range", {"workbook": ids_workbook(), **action_input})

    assert response["ok"] is False
    assert response["error_code"] == code


def test_read_range_without_an_output_grant_fails_before_graph(monkeypatch) -> None:
    fake = FakeExcel()

    response = run(
        fake, "read_range", {"workbook": ids_workbook(), "worksheet": "Data", "address": "A1"}
    )

    assert response["error_code"] == "dataset_output_unavailable"
    assert fake.calls == []


def test_read_input_rows_validates_the_grant_and_rows(monkeypatch) -> None:
    from excel import datasets

    descriptor = {
        "binding_key": "excel_input_dataset",
        "content_type": "application/x-ndjson",
        "size_bytes": 40,
        "row_count": 2,
        "schema": {"columns": [{"name": "Invoice"}, {"name": "Amount"}]},
        "scope": {"account_id": "acme", "project_id": "p1"},
        "access": {"transport": "presigned_url", "mode": "read"},
    }
    root = {"host": {"scope": {"account_id": "acme", "project_id": "p1"}}}
    monkeypatch.setattr(datasets, "find_dataset_descriptor", lambda *_a, **_k: dict(descriptor))
    monkeypatch.setattr(
        datasets,
        "stream_dataset_bytes",
        lambda *_a, **_k: iter([b'{"Invoice":"INV-1","Amount":1}\n{"Inv', b'oice":"INV-2"}\n']),
    )

    columns, rows = datasets.read_input_rows(root, "dataset:in", max_rows=10)

    assert columns == ["Invoice", "Amount"]
    assert rows == [{"Invoice": "INV-1", "Amount": 1}, {"Invoice": "INV-2"}]
    with pytest.raises(M365Error) as too_many:
        datasets.read_input_rows(root, "dataset:in", max_rows=1)
    assert too_many.value.code == "range_too_large"
    with pytest.raises(M365Error) as other_scope:
        datasets.read_input_rows(
            {"host": {"scope": {"account_id": "x"}}}, "dataset:in", max_rows=10
        )
    assert other_scope.value.code == "dataset_input_unavailable"
    monkeypatch.setattr(
        datasets,
        "stream_dataset_bytes",
        lambda *_a, **_k: iter([b'{"Unknown":"x"}\n{"Invoice":"y"}\n']),
    )
    with pytest.raises(M365Error) as bad_row:
        datasets.read_input_rows(root, "dataset:in", max_rows=10)
    assert bad_row.value.code == "dataset_input_unavailable"
