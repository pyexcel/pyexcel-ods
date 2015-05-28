import os
from pyexcel_ods import get_data


def test_bug_fix_for_issue_1():
    data = get_data(os.path.join("tests", "fixtures", "repeated.ods"))
    print data["Sheet1"]
    assert data['Sheet1'] == [['repeated', 'repeated', 'repeated', 'repeated']]
                    
