"""
    pyexcel_ods.odsbook
    ~~~~~~~~~~~~~~~~~~~

    ODS format plugin for pyexcel

    :copyright: (c) 2014 by C. W.
    :license: GPL v3
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
import datetime
import odf.opendocument
from odf.table import *
from odf.text import P
from odf.namespaces import OFFICENS
import sys
if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


def float_value(value):
    ret = float(value)
    return ret


def date_value(value):
    tokens = value.split('-')
    year = int(tokens[0])
    month = int(tokens[1])
    day = int(tokens[2])
    ret = datetime.date(year, month, day)
    return ret


def ods_date_value(value):
    return value.strftime("%Y-%m-%d")


def time_value(value):
    hour = int(value[2:4])
    minute = int(value[5:7])
    second = int(value[8:10])
    ret = datetime.time(hour, minute, second)
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


ODS_FORMAT_CONVERSION = {
    "float": float,
    "date": datetime.date,
    "time": datetime.time,
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
    bool: "boolean",
    unicode: "string"
}


VALUE_CONVERTERS = {
    "float": float_value,
    "date": date_value,
    "time": time_value,
    "boolean": boolean_value,
    "percentage": float_value,
    "currency": float_value
}

ODS_VALUE_CONVERTERS = {
    "date": ods_date_value,
    "time": ods_time_value,
    "boolean": ods_bool_value
}


VALUE_TOKEN = {
    "float": "value",
    "date": "date-value",
    "time": "time-value",
    "boolean": "boolean-value",
    "percentage": "value",
    "currency": "value"
}


class ODSBook:

    def __init__(self, filename, file_content=None, **keywords):
        """Load the file"""
        if filename:
            self.doc = odf.opendocument.load(filename)
        else:
            self.doc = odf.opendocument.load(file_content)            
        self.SHEETS = OrderedDict()
        self.sheet_names = []
        for sheet in self.doc.spreadsheet.getElementsByType(Table):
            self.readSheet(sheet)

    def readSheet(self, sheet):
        """reads a sheet in the sheet dictionary, storing each sheet
        as an array (rows) of arrays (columns)"""
        name = sheet.getAttribute("name")
        rows = sheet.getElementsByType(TableRow)
        arrRows = []
        # for each row
        for row in rows:
            arrCells = []
            cells = row.getElementsByType(TableCell)
            has_value = False

            # for each cell
            for cell in cells:
                # repeated value?
                repeat = cell.getAttribute("numbercolumnsrepeated")
                if(not repeat):
                    has_value = True
                    ret = self._read_cell(cell)
                    arrCells.append(ret)
                else:
                    r = int(repeat)
                    for i in range(0, r):
                        arrCells.append("")
            # if row contained something
            if(len(arrCells) and has_value):
                arrRows.append(arrCells)

        self.SHEETS[name] = arrRows
        self.sheet_names.append(name)

    def _read_text_cell(self, cell):
        textContent = ""
        ps = cell.getElementsByType(P)
        # for each text node
        for p in ps:
            for n in p.childNodes:
                if (n.nodeType == 3):
                    textContent = textContent + unicode(n.data)
        return textContent

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

    def sheets(self):
        """
        returns a dictionary of all sheet content

        the keys represents sheet names
        the values are two dimensional array
        """
        return self.SHEETS

    def __dict__(self):
        return self.SHEETS


class ODSSheetWriter:
    """
    ODS sheet writer
    """

    def __init__(self, book, name):
        self.doc = book
        if name:
            sheet_name = name
        else:
            sheet_name = "pyexcel_sheet1"
        self.table = Table(name=sheet_name)

    def set_size(self, size):
        pass

    def write_cell(self, row, x):
        tc = TableCell()
        x_type = type(x)
        x_odf_type = ODS_WRITE_FORMAT_COVERSION.get(x_type, "string")
        tc.setAttrNS(OFFICENS, "value-type", x_odf_type)
        x_odf_value_token = VALUE_TOKEN.get(x_odf_type, "value")
        converter = ODS_VALUE_CONVERTERS.get(x_odf_type, None)
        if converter:
            x = converter(x)
        tc.setAttrNS(OFFICENS, x_odf_value_token, x)
        tc.addElement(P(text=x))
        row.addElement(tc)

    def write_row(self, array):
        """
        write a row into the file
        """
        tr = TableRow()
        self.table.addElement(tr)
        for x in array:
            self.write_cell(tr, x)

    def write_array(self, table):
        for r in table:
            self.write_row(r)

    def close(self):
        """
        This call writes file

        """
        self.doc.spreadsheet.addElement(self.table)


class ODSWriter:
    """
    open document spreadsheet writer

    """
    def __init__(self, file):
        from odf.opendocument import OpenDocumentSpreadsheet
        self.doc = OpenDocumentSpreadsheet()
        self.file = file

    def create_sheet(self, name):
        """
        write a row into the file
        """
        return ODSSheetWriter(self.doc, name)

    def write(self, sheet_dicts):
        """Write a dictionary to a multi-sheet file

        Requirements for the dictionary is: key is the sheet name,
        its value must be two dimensional array
        """
        keys = sheet_dicts.keys()
        for name in keys:
            sheet = self.create_sheet(name)
            sheet.write_array(sheet_dicts[name])
            sheet.close()

    def close(self):
        """
        This call writes file

        """
        self.doc.write(self.file)
