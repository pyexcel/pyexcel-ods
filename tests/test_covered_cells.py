from io import BytesIO

import pytest
from odf.text import P
from odf.table import Table, TableRow, TableCell, CoveredTableCell
from pyexcel_ods import get_data
from odf.opendocument import OpenDocumentSpreadsheet


def text_cell(value, **attributes):
    cell = TableCell(valuetype="string", **attributes)
    cell.addElement(P(text=value))
    return cell


def read_rows(rows, **options):
    document = OpenDocumentSpreadsheet()
    table = Table(name="Merged")
    document.spreadsheet.addElement(table)
    for cells in rows:
        row = TableRow()
        for cell in cells:
            row.addElement(cell)
        table.addElement(row)
    stream = BytesIO()
    document.save(stream)
    stream.seek(0)
    return get_data(stream, **options)["Merged"]


@pytest.mark.parametrize("repeat", [False, True])
def test_horizontal_merge_keeps_following_column(repeat):
    covered = (
        [CoveredTableCell(numbercolumnsrepeated=2)]
        if repeat
        else [CoveredTableCell(), CoveredTableCell()]
    )
    cells = [text_cell("merged", numbercolumnsspanned=3)]
    cells += covered + [text_cell("fourth")]
    assert read_rows([cells]) == [["merged", "", "", "fourth"]]


def test_vertical_merge_keeps_following_column():
    rows = [
        [text_cell("merged", numberrowsspanned=2), text_cell("first")],
        [CoveredTableCell(), text_cell("second")],
    ]
    assert read_rows(rows) == [["merged", "first"], ["", "second"]]


def test_column_selection_counts_covered_cells():
    cells = [
        text_cell("merged", numbercolumnsspanned=2),
        CoveredTableCell(),
        text_cell("third"),
    ]
    assert read_rows([cells], start_column=2) == [["third"]]
