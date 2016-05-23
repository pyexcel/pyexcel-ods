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


def float_value(value):
    ret = float(value)
    return ret


def date_value(value):
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
    except:
        pass
    if ret == "invalid":
        raise Exception("Bad date value %s" % value)
    return ret


def ods_date_value(value):
    return value.strftime("%Y-%m-%d")


def time_value(value):
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
    if value == "true":
        ret = True
    else:
        ret = False
    return ret


def ods_bool_value(value):
    if value is True:
        return "true"
    else:
        return "false"


def ods_timedelta_value(cell):
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
                if(not repeat):
                    tmp_row.append(cell_value)
                else:
                    r = int(repeat)
                    for i in range(0, r):
                        tmp_row.append(cell_value)
                if cell_value is not None and cell_value != '':
                    arr_cells += tmp_row
                    tmp_row = []
            # if row contained something
            yield arr_cells

    def _read_text_cell(self, cell):
        textContent = []
        ps = cell.getElementsByType(P)
        # for each text node
        for p in ps:
            for n in p.childNodes:
                if (n.nodeType == 3):
                    textContent.append(unicode(n.data))
        return '\n'.join(textContent)

    def _read_cell(self, cell):
        cell_type = cell.getAttrNS(OFFICENS, "value-type")
        value_token = VALUE_TOKEN.get(cell_type, "value")
        ret = None
        if cell_type == "string":
            textContent = self._read_text_cell(cell)
            ret = textContent
        else:
            if cell_type in ODS_FORMAT_CONVERSION:
                value = cell.getAttrNS(OFFICENS, value_token)
                n_value = VALUE_CONVERTERS[cell_type](value)
                ret = n_value
            else:
                textContent = self._read_text_cell(cell)
                ret = textContent
        return ret


class ODSBook(BookReader):

    def open(self, file_name, **keywords):
        BookReader.open(self, file_name, **keywords)
        self._load_from_file()

    def open_stream(self, file_stream, **keywords):
        BookReader.open_stream(self, file_stream, **keywords)
        self._load_from_memory()

    def read_sheet_by_name(self, sheet_name):
        tables = self.native_book.spreadsheet.getElementsByType(Table)
        rets = [table for table in tables
                if table.getAttribute('name') == sheet_name]
        if len(rets) == 0:
            raise ValueError("%s cannot be found" % sheet_name)
        else:
            return self.read_sheet(rets[0])

    def read_sheet_by_index(self, sheet_index):
        tables = self.native_book.spreadsheet.getElementsByType(Table)
        length = len(tables)
        if sheet_index < length:
            return self.read_sheet(tables[sheet_index])
        else:
            raise IndexError("Index %d of out bound %d" % (
                sheet_index, length))

    def read_all(self):
        result = OrderedDict()
        for sheet in self.native_book.spreadsheet.getElementsByType(Table):
            ods_sheet = ODSSheet(sheet)
            result[ods_sheet.name] = ods_sheet.to_array()

        return result
            
    def read_sheet(self, native_sheet):
        sheet = ODSSheet(native_sheet)
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
        self.native_sheet = Table(name=name)

    def set_size(self, size):
        pass

    def write_cell(self, row, cell):
        tc = TableCell()
        cell_type = type(cell)
        cell_odf_type = ODS_WRITE_FORMAT_COVERSION.get(cell_type, "string")
        tc.setAttrNS(OFFICENS, "value-type", cell_odf_type)
        cell_odf_value_token = VALUE_TOKEN.get(cell_odf_type, "value")
        converter = ODS_VALUE_CONVERTERS.get(cell_odf_type, None)
        if converter:
            cell = converter(cell)
        if cell_odf_type != 'string':
            tc.setAttrNS(OFFICENS, cell_odf_value_token, cell)
            tc.addElement(P(text=cell))
        else:
            lines = cell.split('\n')
            for line in lines:
                tc.addElement(P(text=line))
        row.addElement(tc)

    def write_row(self, array):
        """
        write a row into the file
        """
        tr = TableRow()
        self.native_sheet.addElement(tr)
        for cell in array:
            self.write_cell(tr, cell)

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

    def open(self, file_name, **keywords):
        BookWriter.open(self, file_name, **keywords)

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
