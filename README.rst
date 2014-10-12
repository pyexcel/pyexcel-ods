===========
pyexcel-ods
===========

.. image:: https://api.travis-ci.org/chfw/pyexcel-ods.png
    :target: http://travis-ci.org/chfw/pyexcel-ods

.. image:: https://codecov.io/github/chfw/pyexcel-ods/coverage.png
    :target: https://codecov.io/github/chfw/pyexcel-ods

**pyexcel-ods** is a plugin to `pyexcel <https://github.com/chfw/pyexcel>`_ and provides the capbility to read, manipulate and write data in ods fromat using python 2.6 and python 2.7

Usage
=====

Import it in your file to enable this plugin::

    from pyexcel.ext import ods

Reading from an ods file
------------------------

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
----------------------

Here is the sample code::

    from pyexcel import Writer
    from pyexcel.ext import ods
    
    array = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    writer = Writer("output.ods")
    writer.write_array(array)
    writer.close()


Dependencies
============

1. pyexcel >= 0.0.4
2. odfpy

Credits
=======

ODSReader is written by [Marco Conti](https://github.com/marcoconti83/read-ods-with-odfpy)
