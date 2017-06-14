"""
    pyexcel_ods.odsw
    ~~~~~~~~~~~~~~~~~~~~~

    ods writer

    :copyright: (c) 2014-2017 by Onni Software Ltd.
    :license: New BSD License, see LICENSE for more details
"""
import sys

from odf.table import TableRow, TableCell, Table
from odf.text import P
from odf.namespaces import OFFICENS
from odf.opendocument import OpenDocumentSpreadsheet

from pyexcel_io.book import BookWriter
from pyexcel_io.sheet import SheetWriter

import pyexcel_ods.converter as converter

PY2 = sys.version_info[0] == 2

PY27_BELOW = PY2 and sys.version_info[1] < 7


class ODSSheetWriter(SheetWriter):
    """
    ODS sheet writer
    """
    def set_sheet_name(self, name):
        """initialize the native table"""
        self._native_sheet = Table(name=name)

    def set_size(self, size):
        """not used in this class but used in ods3"""
        pass

    def write_cell(self, row, cell):
        """write a native cell"""
        cell_to_be_written = TableCell()
        cell_type = type(cell)
        cell_odf_type = converter.ODS_WRITE_FORMAT_COVERSION.get(
            cell_type, "string")
        cell_to_be_written.setAttrNS(OFFICENS, "value-type", cell_odf_type)
        cell_odf_value_token = converter.VALUE_TOKEN.get(
            cell_odf_type, "value")
        converter_func = converter.ODS_VALUE_CONVERTERS.get(
            cell_odf_type, None)
        if converter_func:
            cell = converter_func(cell)
        if cell_odf_type != 'string':
            cell_to_be_written.setAttrNS(OFFICENS, cell_odf_value_token, cell)
            cell_to_be_written.addElement(P(text=cell))
        else:
            lines = cell.split('\n')
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


class ODSWriter(BookWriter):
    """
    open document spreadsheet writer

    """
    def __init__(self):
        BookWriter.__init__(self)
        self._native_book = OpenDocumentSpreadsheet()

    def create_sheet(self, name):
        """
        write a row into the file
        """
        return ODSSheetWriter(self._native_book, None, name)

    def close(self):
        """
        This call writes file

        """
        self._native_book.write(self._file_alike_object)
        self._native_book = None
