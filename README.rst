================================================================================
pyexcel-ods - Let you focus on data, instead of ods format
================================================================================

.. image:: https://api.travis-ci.org/chfw/pyexcel-ods.png
    :target: http://travis-ci.org/chfw/pyexcel-ods

.. image:: https://codecov.io/github/chfw/pyexcel-ods/coverage.png
    :target: https://codecov.io/github/chfw/pyexcel-ods


**pyexcel-ods** is a tiny wrapper library to read, manipulate and write data in
ods fromat using python 2.6 and python 2.7. You are likely to use it with
`pyexcel <https://github.com/chfw/pyexcel>`_.
`pyexcel-ods3 <https://github.com/chfw/pyexcel-ods3>`_ is a sister library that
does the same thing but supports Python 3.3 and 3.4 and depends on lxml.

Known constraints
==================

Fonts, colors and charts are not supported. 

Installation
============

You can install it via pip::

    $ pip install pyexcel-ods


or clone it and install it::

    $ git clone http://github.com/chfw/pyexcel-ods.git
    $ cd pyexcel-ods
    $ python setup.py install

Usage
=====

As a standalone library
------------------------

.. testcode::
   :hide:

    >>> import sys
    >>> if sys.version_info[0] < 3:
    ...     from StringIO import StringIO
    ... else:
    ...     from io import BytesIO as StringIO
    >>> from pyexcel_io import OrderedDict


Write to an ods file
*********************

Here's the sample code to write a dictionary to an ods file::

    >>> from pyexcel_ods import save_data
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.ods", data)

Read from an ods file
**********************

Here's the sample code::

    >>> from pyexcel_ods import get_data
    >>> data = get_data("your_file.ods")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an ods file to memory
******************************

Here's the sample code to write a dictionary to an ods file::

    >>> from pyexcel_ods import save_data
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> save_data(io, data)
    >>> # do something with the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading


Read from an ods from memory
*****************************

Here's the sample code::

    >>> # This is just an illustration
    >>> # In reality, you might deal with ods file upload
    >>> # where you will read from requests.FILES['YOUR_ODS_FILE']
    >>> data = get_data(io)
    >>> print(json.dumps(data))
    {"Sheet 1": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], "Sheet 2": [[7.0, 8.0, 9.0], [10.0, 11.0, 12.0]]}


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin::

    from pyexcel.ext import ods

Please note only pyexcel version 0.0.4+ support this.

Reading from an ods file
************************

Here is the sample code::

    >>> import pyexcel as pe
    >>> from pyexcel.ext import ods
    >>> sheet = pe.get_book(file_name="your_file.ods")
    >>> sheet
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+

Writing to an ods file
**********************

Here is the sample code::

    >>> sheet.save_as("another_file.ods")

Reading from a IO instance
================================

You got to wrap the binary content with StringIO to get odf working::


    >>> # This is just an illustration
    >>> # In reality, you might deal with ods file upload
    >>> # where you will read from requests.FILES['YOUR_ODS_FILE']
    >>> odsfile = "another_file.ods"
    >>> with open(odsfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.get_book(file_type="ods", file_content=content)
    ...     print(r)
    ...
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to a StringIO instance
================================

You need to pass a StringIO instance to Writer::

    >>> data = [
    ...     [1, 2, 3],
    ...     [4, 5, 6]
    ... ]
    >>> io = StringIO()
    >>> sheet = pe.Sheet(data)
    >>> sheet.save_to_memory("ods", io)
    >>> # then do something with io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

License
=========

New BSD License


Dependencies
============

1. odfpy == 0.9.6
2. pyexcel-io >= 0.0.4

Credits
=======

ODSReader is originally written by `Marco Conti <https://github.com/marcoconti83/read-ods-with-odfpy>`_

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.ods")
   >>> os.unlink("another_file.ods")
