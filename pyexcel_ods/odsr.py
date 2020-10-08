"""
    pyexcel_ods.odsr
    ~~~~~~~~~~~~~~~~~~~~~

    ods reader

    :copyright: (c) 2014-2020 by Onni Software Ltd.
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
from io import BytesIO

import pyexcel_io.service as service
from odf.text import P
from odf.table import Table, TableRow, TableCell

# Thanks to grt for the fixes
from odf.teletype import extractText
from odf.namespaces import OFFICENS
from odf.opendocument import load
from pyexcel_io.plugin_api import ISheet, IReader, NamedContent


class ODSSheet(ISheet):
    """native ods sheet"""

    def __init__(self, sheet, auto_detect_int=True, **keywords):
        self._native_sheet = sheet
        self._keywords = keywords
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

    def __read_cell(self, cell):
        cell_type = cell.getAttrNS(OFFICENS, "value-type")
        value_token = service.VALUE_TOKEN.get(cell_type, "value")
        ret = None
        if cell_type == "string":
            text_content = self.__read_text_cell(cell)
            ret = text_content
        elif cell_type == "currency":
            value = cell.getAttrNS(OFFICENS, value_token)
            currency = cell.getAttrNS(OFFICENS, cell_type)
            if currency:
                ret = value + " " + currency
            else:
                ret = value
        else:
            if cell_type in service.VALUE_CONVERTERS:
                value = cell.getAttrNS(OFFICENS, value_token)
                n_value = service.VALUE_CONVERTERS[cell_type](value)
                if cell_type == "float" and self.__auto_detect_int:
                    if service.has_no_digits_in_float(n_value):
                        n_value = int(n_value)
                ret = n_value
            else:
                text_content = self.__read_text_cell(cell)
                ret = text_content
        return ret

    def __read_text_cell(self, cell):
        text_content = []
        paragraphs = cell.getElementsByType(P)
        # for each text node
        for paragraph in paragraphs:
            name_space, tag = paragraph.parentNode.qname
            if tag != str("annotation"):
                data = extractText(paragraph)
                text_content.append(data)
        return "\n".join(text_content)


class ODSBook(IReader):
    """read ods book"""

    def __init__(self, file_alike_object, _, **keywords):
        self._native_book = load(file_alike_object)
        self._keywords = keywords
        self.content_array = [
            NamedContent(table.getAttribute("name"), table)
            for table in self._native_book.spreadsheet.getElementsByType(Table)
        ]

    def read_sheet(self, sheet_index):
        """read a sheet at a specified index"""
        table = self.content_array[sheet_index].payload
        sheet = ODSSheet(table, **self._keywords)
        return sheet

    def close(self):
        self._native_book = None


class ODSBookInContent(ODSBook):
    """
    Open xlsx as read only mode
    """

    def __init__(self, file_content, file_type, **keywords):
        io = BytesIO(file_content)
        super().__init__(io, file_type, **keywords)
