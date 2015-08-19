#!/usr/bin/python
# -*- encoding: utf-8 -*-
import os
from pyexcel_ods import get_data, save_data


def test_bug_fix_for_issue_1():
    data = get_data(os.path.join("tests", "fixtures", "repeated.ods"))
    assert data['Sheet1'] == [['repeated', 'repeated', 'repeated', 'repeated']]
                    
def test_bug_fix_for_issue_2():
    data = {}
    data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    data.update({"Sheet 2": [[u"row 1", u"Héllô!", u"HolÁ!"]]})
    save_data("your_file.ods", data)
    new_data = get_data("your_file.ods")
    assert new_data["Sheet 2"] == [[u'row 1', u'H\xe9ll\xf4!', u'Hol\xc1!']]