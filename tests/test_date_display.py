import datetime
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree

import pytest
from pyexcel_ods import get_data, save_data
from odf.namespaces import STYLENS, TABLENS, NUMBERNS, OFFICENS


def tag(namespace, name):
    return "{%s}%s" % (namespace, name)


@pytest.mark.parametrize("source", ["file", "stream"])
def test_dates_have_display_styles(tmp_path, source):
    dates = [datetime.date(2026, 10, 27), datetime.date(2000, 2, 29)]
    data = {"Dates": [dates, [42, "plain"]], "Other": [[dates[0]]]}
    if source == "file":
        output = str(tmp_path / "dates.ods")
        save_data(output, data)
        archive = ZipFile(output)
    else:
        output = BytesIO()
        save_data(output, data)
        output.seek(0)
        archive = ZipFile(output)
    with archive:
        root = ElementTree.fromstring(archive.read("content.xml"))
    styles = {
        node.get(tag(STYLENS, "name")): node
        for node in root.findall(".//" + tag(STYLENS, "style"))
    }
    formats = {
        node.get(tag(STYLENS, "name")): node
        for node in root.findall(".//" + tag(NUMBERNS, "date-style"))
    }
    assert len(styles) == len(formats) == 1
    cells = root.findall(".//" + tag(TABLENS, "table-cell"))
    for cell in (cells[0], cells[1], cells[4]):
        style = styles[cell.get(tag(TABLENS, "style-name"))]
        date_format = formats[style.get(tag(STYLENS, "data-style-name"))]
        assert [child.tag for child in date_format] == [
            tag(NUMBERNS, "year"),
            tag(NUMBERNS, "text"),
            tag(NUMBERNS, "month"),
            tag(NUMBERNS, "text"),
            tag(NUMBERNS, "day"),
        ]
        assert cell.get(tag(OFFICENS, "value-type")) == "date"
        assert cell.get(tag(OFFICENS, "date-value")) in {
            date.isoformat() for date in dates
        }
    for cell in cells[2:4]:
        assert tag(TABLENS, "style-name") not in cell.attrib
    if source == "stream":
        output.seek(0)
    assert get_data(output) == data


def test_no_date_styles_without_dates():
    output = BytesIO()
    save_data(output, {"Data": [[42, "plain"]]})
    output.seek(0)
    with ZipFile(output) as archive:
        root = ElementTree.fromstring(archive.read("content.xml"))
    assert root.findall(".//" + tag(NUMBERNS, "date-style")) == []
    assert root.findall(".//" + tag(STYLENS, "style")) == []


def test_date_styles_reused_by_interleaved_sheet_writers():
    from odf.style import Style
    from odf.number import DateStyle
    from odf.opendocument import OpenDocumentSpreadsheet
    from pyexcel_ods.odsw import ODSSheetWriter

    book = OpenDocumentSpreadsheet()
    first = ODSSheetWriter(book, "First")
    second = ODSSheetWriter(book, "Second")
    first.write_row([datetime.date(2026, 10, 27)])
    second.write_row([datetime.date(2000, 2, 29)])
    first.close()
    second.close()
    assert len(book.automaticstyles.getElementsByType(DateStyle)) == 1
    assert len(book.automaticstyles.getElementsByType(Style)) == 1
