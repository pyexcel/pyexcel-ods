#!/usr/bin/python
# -*- encoding: utf-8 -*-
import os
import psutil
import pyexcel as pe
from pyexcel_ods import get_data, save_data
from nose.tools import raises, eq_


def test_bug_fix_for_issue_1():
    data = get_data(get_fixtures("repeated.ods"))
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
    data = get_data(get_fixtures(test_file),
                    skip_empty_rows=True)
    eq_(data['S-LMC'], [[u'aaa'], [0]])


def test_issue_6():
    test_file = "12_day_as_time.ods"
    data = get_data(get_fixtures(test_file),
                    skip_empty_rows=True)
    eq_(data['Sheet1'][0][0].days, 12)


def test_issue_19():
    test_file = "pyexcel_81_ods_19.ods"
    data = get_data(get_fixtures(test_file),
                    skip_empty_rows=True)
    eq_(data['product.template'][1][1], 'PRODUCT NAME PMP')


def test_issue_83_ods_file_handle():
    # this proves that odfpy
    # does not leave a file handle open at all
    proc = psutil.Process()
    test_file = get_fixtures("issue_61.ods")
    open_files_l1 = proc.open_files()

    # start with a csv file
    data = pe.iget_array(file_name=test_file, library='pyexcel-ods')
    open_files_l2 = proc.open_files()
    delta = len(open_files_l2) - len(open_files_l1)
    # cannot catch open file handle
    assert delta == 0

    # now the file handle get opened when we run through
    # the generator
    list(data)
    open_files_l3 = proc.open_files()
    delta = len(open_files_l3) - len(open_files_l1)
    # cannot catch open file handle
    assert delta == 0

    # free the fish
    pe.free_resources()
    open_files_l4 = proc.open_files()
    # this confirms that no more open file handle
    eq_(open_files_l1, open_files_l4)


def test_pr_22():
    test_file = get_fixtures("white_space.ods")
    data = get_data(test_file)
    # OrderedDict([(u'Sheet1', [[u'paragraph with tab,  space, new line']])])
    eq_(data['Sheet1'][0][0], 'paragraph with tab(\t),    space, \nnew line')


def get_fixtures(filename):
    return os.path.join("tests", "fixtures", filename)
