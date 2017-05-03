# Copyright 2011 Marco Conti

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

#   http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Thanks to grt for the fixes
import sys
import math

from odf.table import TableRow, TableCell, Table
from odf.text import P
from odf.namespaces import OFFICENS
from odf.opendocument import OpenDocumentSpreadsheet, load

from pyexcel_io.book import BookReader, BookWriter
from pyexcel_io.sheet import SheetReader, SheetWriter

import pyexcel_ods.converter as converter

PY2 = sys.version_info[0] == 2

PY27_BELOW = PY2 and sys.version_info[1] < 7
if PY27_BELOW:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


class ODSSheet(SheetReader):
    """native ods sheet"""
    def __init__(self, sheet, auto_detect_int=True, **keywords):
        SheetReader.__init__(self, sheet, **keywords)
        self.__auto_detect_int = auto_detect_int

    @property
    def name(self):
        return self._native_sheet.getAttribute("name")

    def row_iterator(self):
        return self._native_sheet.getElementsByType(TableRow)

    def column_iterator(self, row):
        cells = row.getElementsByType(TableCell)
        for cell in cells:
            repeat = cell.getAttribute("numbercolumnsrepeated")
            cell_value = self.__read_cell(cell)
            if repeat:
                number_of_repeat = int(repeat)
                for i in range(number_of_repeat):
                    yield cell_value
            else:
                yield cell_value

    def __read_row(self, cells):
        tmp_row = []
        for cell in cells:
            # repeated value?
            repeat = cell.getAttribute("numbercolumnsrepeated")
            cell_value = self.__read_cell(cell)
            if repeat:
                number_of_repeat = int(repeat)
                tmp_row += [cell_value] * number_of_repeat
            else:
                tmp_row.append(cell_value)
        return tmp_row

    def __read_text_cell(self, cell):
        text_content = []
        paragraphs = cell.getElementsByType(P)
        # for each text node
        for paragraph in paragraphs:
            data = ''
            for node in paragraph.childNodes:
                if (node.nodeType == 3):
                    if PY2:
                        data += unicode(node.data)
                    else:
                        data += node.data
            text_content.append(data)
        return '\n'.join(text_content)

    def __read_cell(self, cell):
        cell_type = cell.getAttrNS(OFFICENS, "value-type")
        value_token = converter.VALUE_TOKEN.get(cell_type, "value")
        ret = None
        if cell_type == "string":
            text_content = self.__read_text_cell(cell)
            ret = text_content
        elif cell_type == "currency":
            value = cell.getAttrNS(OFFICENS, value_token)
            currency = cell.getAttrNS(OFFICENS, cell_type)
            ret = value + ' ' + currency
        else:
            if cell_type in converter.VALUE_CONVERTERS:
                value = cell.getAttrNS(OFFICENS, value_token)
                n_value = converter.VALUE_CONVERTERS[cell_type](value)
                if cell_type == 'float' and self.__auto_detect_int:
                    if is_integer_ok_for_xl_float(n_value):
                        n_value = int(n_value)
                ret = n_value
            else:
                text_content = self.__read_text_cell(cell)
                ret = text_content
        return ret


class ODSBook(BookReader):
    """read ods book"""

    def open(self, file_name, **keywords):
        """open ods file"""
        BookReader.open(self, file_name, **keywords)
        self._load_from_file()

    def open_stream(self, file_stream, **keywords):
        """open ods file stream"""
        BookReader.open_stream(self, file_stream, **keywords)
        self._load_from_memory()

    def read_sheet_by_name(self, sheet_name):
        """read a named sheet"""
        tables = self._native_book.spreadsheet.getElementsByType(Table)
        rets = [table for table in tables
                if table.getAttribute('name') == sheet_name]
        if len(rets) == 0:
            raise ValueError("%s cannot be found" % sheet_name)
        else:
            return self.read_sheet(rets[0])

    def read_sheet_by_index(self, sheet_index):
        """read a sheet at a specified index"""
        tables = self._native_book.spreadsheet.getElementsByType(Table)
        length = len(tables)
        if sheet_index < length:
            return self.read_sheet(tables[sheet_index])
        else:
            raise IndexError("Index %d of out bound %d" % (
                sheet_index, length))

    def read_all(self):
        """read all sheets"""
        result = OrderedDict()
        for sheet in self._native_book.spreadsheet.getElementsByType(Table):
            ods_sheet = ODSSheet(sheet, **self._keywords)
            result[ods_sheet.name] = ods_sheet.to_array()

        return result

    def read_sheet(self, native_sheet):
        """read one native sheet"""
        sheet = ODSSheet(native_sheet, **self._keywords)
        return {sheet.name: sheet.to_array()}

    def _load_from_memory(self):
        self._native_book = load(self._file_stream)

    def _load_from_file(self):
        self._native_book = load(self._file_name)


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


def is_integer_ok_for_xl_float(value):
    """check if a float had zero value in digits"""
    return value == math.floor(value)


_ods_registry = {
    "file_type": "ods",
    "reader": ODSBook,
    "writer": ODSWriter,
    "stream_type": "binary",
    "mime_type": "application/vnd.oasis.opendocument.spreadsheet",
    "library": "pyexcel-ods"
}

exports = (_ods_registry,)
