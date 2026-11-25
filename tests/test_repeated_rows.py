from io import BytesIO

import pytest
from odf.text import P
from odf.table import Table, TableRow, TableCell
from pyexcel_ods import get_data
from odf.opendocument import OpenDocumentSpreadsheet


def repeated_rows_book(rows=None):
    book = OpenDocumentSpreadsheet()
    table = Table(name="Repeated rows")
    book.spreadsheet.addElement(table)
    if rows is None:
        rows = [("first", None), ("repeat", 3), ("last", None)]
    for value, repeat in rows:
        row = TableRow()
        if repeat is not None:
            row.setAttribute("numberrowsrepeated", repeat)
        cell = TableCell(valuetype="string")
        cell.addElement(P(text=value))
        row.addElement(cell)
        table.addElement(row)
    stream = BytesIO()
    book.save(stream)
    stream.seek(0)
    return stream


@pytest.mark.parametrize(
    "options, expected",
    [
        ({}, [["first"], ["repeat"], ["repeat"], ["repeat"], ["last"]]),
        ({"start_row": 2, "row_limit": 2}, [["repeat"], ["repeat"]]),
        ({"start_row": 4}, [["last"]]),
    ],
)
def test_read_repeated_rows(options, expected):
    data = get_data(repeated_rows_book(), **options)
    assert data["Repeated rows"] == expected


def test_repeated_rows_are_independent_lists():
    rows = get_data(repeated_rows_book())["Repeated rows"]
    rows[1][0] = "changed"
    assert rows[2] == ["repeat"]


def test_repeated_empty_rows_between_values():
    source = repeated_rows_book([("first", None), ("", 3), ("last", None)])
    rows = get_data(source, skip_empty_rows=False)["Repeated rows"]
    assert rows == [["first"], [], [], [], ["last"]]


def test_trailing_empty_ranges_are_not_expanded():
    source = repeated_rows_book([("value", 2), ("", 1048573), ("", None)])
    rows = get_data(source, skip_empty_rows=False)["Repeated rows"]
    assert rows == [["value"], ["value"], [], []]
