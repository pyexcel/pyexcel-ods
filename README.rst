===========
pyexcel-ods
===========

.. image:: https://api.travis-ci.org/chfw/pyexcel-ods.png
    :target: http://travis-ci.org/chfw/pyexcel-ods

.. image:: https://codecov.io/github/chfw/pyexcel-ods/coverage.png
    :target: https://codecov.io/github/chfw/pyexcel-ods

.. image:: https://pypip.in/d/pyexcel-ods/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-ods

.. image:: https://pypip.in/py_versions/pyexcel-ods/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-ods

.. image:: https://pypip.in/implementation/pyexcel-ods/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-ods

**pyexcel-ods** is a tiny wrapper library to read, manipulate and write data in ods fromat using python 2.6 and python 2.7. You are likely to use it with `pyexcel <https://github.com/chfw/pyexcel>`_. `pyexcel-ods3 <https://github.com/chfw/pyexcel-ods3>`_ is a sister library that does the same thing but supports Python 3.3 and 3.4.

Installation
============

You can install it via pip::

    $ pip install pyexcel-ods


or clone it and install it::

    $ git clone http://github.com/chfw/pyexcel-ods.git
    $ cd pyexcel
    $ python setup.py install

Usage
=====

As a standalone library
------------------------

Read from an ods file
**********************

Here's the sample code::

    from pyexcel_ods import ODSBook

    book = ODSBook("your_file.ods")
    # book.sheets() returns a dictionary of all sheet content
    #   the keys represents sheet names
    #   the values are two dimensional array
    print(book.sheets())

Write to an ods file
*********************

Here's the sample code to write a dictionary to an ods file::

    from pyexcel_ods import ODSWriter

    data = {
        "Sheet 1": [[1, 2, 3], [4, 5, 6]],
        "Sheet 2": [["row 1", "row 2", "row 3"]]
    }
    writer = ODSWriter("your_file.ods")
    writer.write(data)
    writer.close()

Read from an ods from memory
*****************************

Here's the sample code::

    from pyexcel_ods import ODSBook
    from StringIO import StringIO

    # This is just an illustration
    # In reality, you might deal with ods file upload
    # where you will read from requests.FILES['YOUR_ODS_FILE']
    odsfile = "example.ods"
    with open(odsfile, "rb") as f:
        content = f.read()
        book = ODSBook(None, StringIO(content))
        print(book.sheets())


Write an ods to memory
**********************

Here's the sample code to write a dictionary to an ods file::

    from pyexcel_ods import ODSWriter
    from StringIO import StringIO

    data = {
        "Sheet 1": [[1, 2, 3], [4, 5, 6]],
        "Sheet 2": [["row 1", "row 2", "row 3"]]
    }
    io = StringIO()
    writer = ODSWriter(io)
    writer.write(data)
    writer.close()
    # do something witht the io
    # In reality, you might give it to your http response
    # object for downloading

As a pyexcel plugin
--------------------

Import it in your file to enable this plugin::

    from pyexcel.ext import ods

Please note only pyexcel version 0.0.4+ support this.

Reading from an ods file
************************

Here is the sample code::

    from pyexcel import Reader
    from pyexcel.ext import ods
    from pyexcel.utils import to_array
    import json
    
    # "example.ods"
    reader = Reader("example.ods")
    data = to_array(reader)
    print json.dumps(data)

Writing to an ods file
***********************

Here is the sample code::

    from pyexcel import Writer
    from pyexcel.ext import ods
    
    array = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    writer = Writer("output.ods")
    writer.write_array(array)
    writer.close()

Reading from a StringIO instance
================================

You got to wrap the binary content with StringIO to get odf working::


    import pyexcel
    from pyexcel.ext import ods
    from StringIO import StringIO

    # This is just an illustration
    # In reality, you might deal with ods file upload
    # where you will read from requests.FILES['YOUR_ODS_FILE']
    odsfile = "example.ods"
    with open(odsfile, "rb") as f:
        content = f.read()
        r = pyexcel.Reader(("ods", StringIO(content)))


Writing to a StringIO instance
================================

You need to pass a StringIO instance to Writer::

    import pyexcel
    from pyexcel.ext import ods
    from StringIO import StringIO


    data = [
        [1, 2, 3],
        [4, 5, 6]
    ]
    io = StringIO()
    w = pyexcel.Writer(("ods",io))
    w.write_rows(data)
    w.close()
    # then do something with io
    # In reality, you might give it to your http response
    # object for downloading


Dependencies
============

1. odfpy

Credits
=======

ODSReader is originally written by `Marco Conti <https://github.com/marcoconti83/read-ods-with-odfpy>`_
