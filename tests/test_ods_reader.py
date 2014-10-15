import os
from pyexcel.ext import ods
from base import ODSCellTypes


class TestODSReader(ODSCellTypes):
    def setUp(self):
        r = ods.odsbook.ODSBook(os.path.join("tests",
                                             "fixtures",
                                             "ods_formats.ods"))
        self.data = r.sheets()

class TestODSWriter(ODSCellTypes):
    def setUp(self):
        r = ods.odsbook.ODSBook(os.path.join("tests",
                                             "fixtures",
                                             "ods_formats.ods"))
        self.data1 = r.sheets()
        self.testfile = "odswriter.ods"
        w = ods.odsbook.ODSWriter(self.testfile)
        w.write(self.data1)
        w.close()
        r2 = ods.odsbook.ODSBook(self.testfile)
        self.data = r2.sheets()

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
