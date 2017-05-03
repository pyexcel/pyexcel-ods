#!/usr/bin/python
# -*- encoding: utf-8 -*-
import os
from pyexcel_ods import get_data, save_data
from nose.tools import raises, eq_


def test_bug_fix_for_issue_1():
    data = get_data(os.path.join("tests", "fixtures", "repeated.ods"))
    eq_(data["Sheet1"], [['repeated', 'repeated', 'repeated', 'repeated']])


def test_bug_fix_for_issue_2():
    data = {}
    data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    data.update({"Sheet 2": [[u"row 1", u"Héllô!", u"HolÁ!"]]})
    save_data("your_file.ods", data)
    new_data = get_data("your_file.ods")
    assert new_data["Sheet 2"] == [[u'row 1', u'H\xe9ll\xf4!', u'Hol\xc1!']]


def test_date_util_parse():
    from pyexcel_ods.converter import date_value
    value = "2015-08-17T19:20:00"
    d = date_value(value)
    assert d.strftime("%Y-%m-%dT%H:%M:%S") == "2015-08-17T19:20:00"
    value = "2015-08-17"
    d = date_value(value)
    assert d.strftime("%Y-%m-%d") == "2015-08-17"
    value = "2015-08-17T19:20:59.999999"
    d = date_value(value)
    assert d.strftime("%Y-%m-%dT%H:%M:%S") == "2015-08-17T19:20:59"
    value = "2015-08-17T19:20:59.99999"
    d = date_value(value)
    assert d.strftime("%Y-%m-%dT%H:%M:%S") == "2015-08-17T19:20:59"
    value = "2015-08-17T19:20:59.999999999999999"
    d = date_value(value)
    assert d.strftime("%Y-%m-%dT%H:%M:%S") == "2015-08-17T19:20:59"


@raises(Exception)
def test_invalid_date():
    from pyexcel_ods.ods import date_value
    value = "2015-08-"
    date_value(value)


@raises(Exception)
def test_fake_date_time_10():
    from pyexcel_ods.ods import date_value
    date_value("1234567890")


@raises(Exception)
def test_fake_date_time_19():
    from pyexcel_ods.ods import date_value
    date_value("1234567890123456789")


@raises(Exception)
def test_fake_date_time_20():
    from pyexcel_ods.ods import date_value
    date_value("12345678901234567890")


def test_issue_13():
    test_file = "test_issue_13.ods"
    data = [
        [1, 2],
        [],
        [],
        [],
        [3, 4]
    ]
    save_data(test_file, {test_file: data})
    written_data = get_data(test_file, skip_empty_rows=False)
    eq_(data, written_data[test_file])
    os.unlink(test_file)


def test_issue_14():
    # pyexcel issue 61
    test_file = "issue_61.ods"
    data = get_data(os.path.join("tests", "fixtures", test_file),
                    skip_empty_rows=True)
    eq_(data['S-LMC'], [[u'aaa'], [0]])


def test_issue_6():
    test_file = "12_day_as_time.ods"
    data = get_data(os.path.join("tests", "fixtures", test_file),
                    skip_empty_rows=True)
    eq_(data['Sheet1'][0][0].days, 12)

def test_issue_19():
    test_file = "pyexcel_81_ods_19.ods"
    data = get_data(os.path.join("tests", "fixtures", test_file),
                    skip_empty_rows=True)
    print(data)
    eq_(data['product.template'][1][1], 'PRODUCT NAME PMP') 
