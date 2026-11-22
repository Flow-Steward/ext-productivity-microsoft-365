"""Excel operations (``m365_excel_user``, delegated only; task B5.4)."""

from __future__ import annotations

from typing import Any

from m365.errors import error_response
from m365.runtime import Operation

from .reads import LIST_OBJECTS_FIELDS, READ_RANGE_FIELDS, list_workbook_objects, read_range
from .search import FIND_FIELDS, find_workbooks
from .writes import APPEND_FIELDS, WRITE_RANGE_FIELDS, append_table_rows, write_range

SERVICE = "excel"

OPERATIONS: dict[str, Operation] = {
    "find_workbooks": Operation("find_workbooks", SERVICE, find_workbooks, fields=FIND_FIELDS),
    "list_workbook_objects": Operation(
        "list_workbook_objects", SERVICE, list_workbook_objects, fields=LIST_OBJECTS_FIELDS
    ),
    "read_range": Operation("read_range", SERVICE, read_range, fields=READ_RANGE_FIELDS),
    "write_range": Operation(
        "write_range", SERVICE, write_range, effectful=True, fields=WRITE_RANGE_FIELDS
    ),
    "append_table_rows": Operation(
        "append_table_rows", SERVICE, append_table_rows, effectful=True, fields=APPEND_FIELDS
    ),
}

#: Excel has no picker in v0.1 (task B5.4), so it serves no resource query.
QUERY_IDS: frozenset[str] = frozenset()


def handle_query(payload: dict[str, Any]) -> dict[str, Any]:
    del payload
    return error_response("invalid_payload", "Unknown Microsoft 365 Excel resource query")


__all__ = ["OPERATIONS", "QUERY_IDS", "handle_query"]
