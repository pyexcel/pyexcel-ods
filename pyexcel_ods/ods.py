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
import datetime
from odf.table import TableRow, TableCell, Table
from odf.text import P
from odf.namespaces import OFFICENS
from odf.opendocument import OpenDocumentSpreadsheet, load

from pyexcel_io.book import BookReader, BookWriter
from pyexcel_io.sheet import SheetReader, SheetWriter

PY27_BELOW = sys.version_info[0] == 2 and sys.version_info[1] < 7
if PY27_BELOW:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


def is_integer_ok_for_xl_float(value):
    """check if a float had zero value in digits"""
    return value == math.floor(value)


def float_value(value):
    """convert a value to float"""
    ret = float(value)
    return ret


def date_value(value):
    """convert to data value accroding ods specification"""
    ret = "invalid"
    try:
        # catch strptime exceptions only
        if len(value) == 10:
            ret = datetime.datetime.strptime(
                value,
                "%Y-%m-%d")
            ret = ret.date()
        elif len(value) == 19:
            ret = datetime.datetime.strptime(
                value,
                "%Y-%m-%dT%H:%M:%S")
        elif len(value) > 19:
            ret = datetime.datetime.strptime(
                value[0:26],
                "%Y-%m-%dT%H:%M:%S.%f")
    except ValueError:
        pass
    if ret == "invalid":
        raise Exception("Bad date value %s" % value)
    return ret


def ods_date_value(value):
    return value.strftime("%Y-%m-%d")


def time_value(value):
    """convert to time value accroding the specification"""
    hour = int(value[2:4])
    minute = int(value[5:7])
    second = int(value[8:10])
    if hour < 24:
        ret = datetime.time(hour, minute, second)
    else:
        ret = datetime.timedelta(hours=hour, minutes=minute, seconds=second)
    return ret


def ods_time_value(value):
    return value.strftime("PT%HH%MM%SS")


def boolean_value(value):
    """get bolean value"""
    if value == "true":
        ret = True
    else:
        ret = False
    return ret


def ods_bool_value(value):
    """convert a boolean value to text"""
    if value is True:
        return "true"
    else:
        return "false"


def ods_timedelta_value(cell):
    """convert a cell value to time delta"""
    hours = cell.days * 24 + cell.seconds // 3600
    minutes = (cell.seconds // 60) % 60
    seconds = cell.seconds % 60
    return "PT%02dH%02dM%02dS" % (hours, minutes, seconds)


ODS_FORMAT_CONVERSION = {
    "float": float,
    "date": datetime.date,
    "time": datetime.time,
    'timedelta': datetime.timedelta,
    "boolean": bool,
    "percentage": float,
    "currency": float
}


ODS_WRITE_FORMAT_COVERSION = {
    float: "float",
    int: "float",
    str: "string",
    datetime.date: "date",
    datetime.time: "time",
    datetime.timedelta: "timedelta",
    bool: "boolean",
    unicode: "string"
}


VALUE_CONVERTERS = {
    "float": float_value,
    "date": date_value,
    "time": time_value,
    "timedelta": time_value,
    "boolean": boolean_value,
    "percentage": float_value,
    "currency": float_value
}

ODS_VALUE_CONVERTERS = {
    "date": ods_date_value,
    "time": ods_time_value,
    "boolean": ods_bool_value,
    "timedelta": ods_timedelta_value
}


VALUE_TOKEN = {
    "float": "value",
    "date": "date-value",
    "time": "time-value",
    "boolean": "boolean-value",
    "percentage": "value",
    "currency": "value",
    "timedelta": "time-value"
}


class ODSSheet(SheetReader):
    """native ods sheet"""
    def __init__(self, sheet, auto_detect_int=True, **keywords):
        SheetReader.__init__(self, sheet, **keywords)
        self.auto_detect_int = auto_detect_int

    @property
    def name(self):
        return self.native_sheet.getAttribute("name")

    def to_array(self):
        """reads a sheet in the sheet dictionary, storing each sheet
        as an array (rows) of arrays (columns)"""
        rows = self.native_sheet.getElementsByType(TableRow)
        # for each row
        for row in rows:
            tmp_row = []
            arr_cells = []
            cells = row.getElementsByType(TableCell)

            # for each cell
            for cell in cells:
                # repeated value?
                repeat = cell.getAttribute("numbercolumnsrepeated")
                cell_value = self._read_cell(cell)
                if repeat:
                    number_of_repeat = int(repeat)
                    tmp_row += [cell_value] * number_of_repeat
                else:
                    tmp_row.append(cell_value)
                if cell_value is not None and cell_value != '':
                    arr_cells += tmp_row
                    tmp_row = []
            # if row contained something
            yield arr_cells

    def _read_text_cell(self, cell):
        text_content = []
        paragraphs = cell.getElementsByType(P)
        # for each text node
        for paragraph in paragraphs:
            for node in paragraph.childNodes:
                if (node.nodeType == 3):
                    text_content.append(unicode(node.data))
        return '\n'.join(text_content)

    def _read_cell(self, cell):
        cell_type = cell.getAttrNS(OFFICENS, "value-type")
        value_token = VALUE_TOKEN.get(cell_type, "value")
        ret = None
        if cell_type == "string":
            text_content = self._read_text_cell(cell)
            ret = text_content
        else:
            if cell_type in VALUE_CONVERTERS:
                value = cell.getAttrNS(OFFICENS, value_token)
                n_value = VALUE_CONVERTERS[cell_type](value)
                if cell_type == 'float' and self.auto_detect_int:
                    if is_integer_ok_for_xl_float(n_value):
                        n_value = int(n_value)
                ret = n_value
            else:
                text_content = self._read_text_cell(cell)
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
        tables = self.native_book.spreadsheet.getElementsByType(Table)
        rets = [table for table in tables
                if table.getAttribute('name') == sheet_name]
        if len(rets) == 0:
            raise ValueError("%s cannot be found" % sheet_name)
        else:
            return self.read_sheet(rets[0])

    def read_sheet_by_index(self, sheet_index):
        """read a sheet at a specified index"""
        tables = self.native_book.spreadsheet.getElementsByType(Table)
        length = len(tables)
        if sheet_index < length:
            return self.read_sheet(tables[sheet_index])
        else:
            raise IndexError("Index %d of out bound %d" % (
                sheet_index, length))

    def read_all(self):
        """read all sheets"""
        result = OrderedDict()
        for sheet in self.native_book.spreadsheet.getElementsByType(Table):
            ods_sheet = ODSSheet(sheet, **self.keywords)
            result[ods_sheet.name] = ods_sheet.to_array()

        return result

    def read_sheet(self, native_sheet):
        """read one native sheet"""
        sheet = ODSSheet(native_sheet, **self.keywords)
        return {sheet.name: sheet.to_array()}

    def _load_from_memory(self):
        self.native_book = load(self.file_stream)

    def _load_from_file(self):
        self.native_book = load(self.file_name)


class ODSSheetWriter(SheetWriter):
    """
    ODS sheet writer
    """
    def set_sheet_name(self, name):
        """initialize the native table"""
        self.native_sheet = Table(name=name)

    def set_size(self, size):
        """not used in this class but used in ods3"""
        pass

    def write_cell(self, row, cell):
        """write a native cell"""
        cell_to_be_written = TableCell()
        cell_type = type(cell)
        cell_odf_type = ODS_WRITE_FORMAT_COVERSION.get(cell_type, "string")
        cell_to_be_written.setAttrNS(OFFICENS, "value-type", cell_odf_type)
        cell_odf_value_token = VALUE_TOKEN.get(cell_odf_type, "value")
        converter = ODS_VALUE_CONVERTERS.get(cell_odf_type, None)
        if converter:
            cell = converter(cell)
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
        self.native_sheet.addElement(row)
        for cell in array:
            self.write_cell(row, cell)

    def close(self):
        """
        This call writes file

        """
        self.native_book.spreadsheet.addElement(self.native_sheet)


class ODSWriter(BookWriter):
    """
    open document spreadsheet writer

    """
    def __init__(self):
        BookWriter.__init__(self)
        self.native_book = OpenDocumentSpreadsheet()

    def create_sheet(self, name):
        """
        write a row into the file
        """
        return ODSSheetWriter(self.native_book, None, name)

    def close(self):
        """
        This call writes file

        """
        self.native_book.write(self.file_alike_object)


_ods_registry = {
    "file_type": "ods",
    "reader": ODSBook,
    "writer": ODSWriter,
    "stream_type": "binary",
    "mime_type": "application/vnd.oasis.opendocument.spreadsheet",
    "library": "pyexcel-ods"
}

exports = (_ods_registry,)
