"""
    pyexcel_ods.odsr
    ~~~~~~~~~~~~~~~~~~~~~

    ods reader

    :copyright: (c) 2014-2017 by Onni Software Ltd.
    :license: New BSD License, see LICENSE for more details
"""
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
import math

from odf.table import TableRow, TableCell, Table
from odf.text import P
from odf.namespaces import OFFICENS
from odf.opendocument import load

from pyexcel_io.book import BookReader
from pyexcel_io.sheet import SheetReader
from pyexcel_io._compact import OrderedDict, PY2

import pyexcel_ods.converter as converter


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

    def close(self):
        self._native_book = None

    def _load_from_memory(self):
        self._native_book = load(self._file_stream)

    def _load_from_file(self):
        self._native_book = load(self._file_name)


def is_integer_ok_for_xl_float(value):
    """check if a float had zero value in digits"""
    return value == math.floor(value)
