"""
    pyexcel_ods.odsw
    ~~~~~~~~~~~~~~~~~~~~~

    ods writer

    :copyright: (c) 2014-2020 by Onni Software Ltd.
    :license: New BSD License, see LICENSE for more details
"""

import datetime

import pyexcel_io.service as converter
from odf.text import P
from odf.style import Style
from odf.table import Table, TableRow, TableCell
from odf.number import Day, Text, Year, Month, DateStyle
from odf.namespaces import OFFICENS
from odf.opendocument import OpenDocumentSpreadsheet
from pyexcel_io.plugin_api import IWriter, ISheetWriter


class ODSSheetWriter(ISheetWriter):
    """
    ODS sheet writer
    """

    def __init__(self, ods_book, sheet_name):
        self._native_book = ods_book
        self._native_sheet = Table(name=sheet_name)
        self._date_style = None

    def write_cell(self, row, cell):
        """write a native cell"""
        cell_to_be_written = TableCell()
        if type(cell) is datetime.date:
            if self._date_style is None:
                styles = self._native_book.automaticstyles.getElementsByType(
                    Style
                )
                self._date_style = next(
                    (
                        style
                        for style in styles
                        if style.getAttribute("name") == "pyexcel-date"
                    ),
                    None,
                )
            if self._date_style is None:
                date_format = DateStyle(name="pyexcel-date-format")
                date_format.addElement(Year(style="long"))
                date_format.addElement(Text(text="-"))
                date_format.addElement(Month(style="long"))
                date_format.addElement(Text(text="-"))
                date_format.addElement(Day(style="long"))
                self._native_book.automaticstyles.addElement(date_format)
                self._date_style = Style(
                    name="pyexcel-date",
                    family="table-cell",
                    datastylename=date_format,
                )
                self._native_book.automaticstyles.addElement(self._date_style)
            cell_to_be_written.setAttribute("stylename", self._date_style)
        cell_type = type(cell)
        cell_odf_type = converter.ODS_WRITE_FORMAT_COVERSION.get(
            cell_type, "string"
        )
        cell_to_be_written.setAttrNS(OFFICENS, "value-type", cell_odf_type)
        cell_odf_value_token = converter.VALUE_TOKEN.get(
            cell_odf_type, "value"
        )
        converter_func = converter.ODS_VALUE_CONVERTERS.get(
            cell_odf_type, None
        )
        if converter_func:
            cell = converter_func(cell)
        if cell_odf_type != "string":
            cell_to_be_written.setAttrNS(OFFICENS, cell_odf_value_token, cell)
            cell_to_be_written.addElement(P(text=cell))
        else:
            lines = cell.split("\n")
            for line in lines:
                cell_to_be_written.addElement(P(text=line))
        row.addElement(cell_to_be_written)

    def write_row(self, array):
        """
        write a row into the file
        """
        row = TableRow()
        self._native_sheet.addElement(row)
        for cell in array:
            self.write_cell(row, cell)

    def close(self):
        """
        This call writes file

        """
        self._native_book.spreadsheet.addElement(self._native_sheet)


class ODSWriter(IWriter):
    """
    open document spreadsheet writer

    """

    def __init__(self, file_alike_object, file_type, **_):
        self.file_alike_object = file_alike_object
        self._native_book = OpenDocumentSpreadsheet()

    def create_sheet(self, name):
        """
        write a row into the file
        """
        return ODSSheetWriter(self._native_book, name)

    def close(self):
        """
        This call writes file

        """
        self._native_book.write(self.file_alike_object)
        self._native_book = None
