Change log
================================================================================

0.5.3 - unreleased
--------------------------------------------------------------------------------

added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. `#24 <https://github.com/pyexcel/pyexcel-ods/issues/24>`_, ignore
   comments(<office:comment>) in cell
#. `#27 <https://github.com/pyexcel/pyexcel-ods/issues/27>`_, exception raised
   when currency type is missing
#. fix odfpy version on 1.3.5.

0.5.2 - 23.10.2017
--------------------------------------------------------------------------------

updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. pyexcel `pyexcel#105 <https://github.com/pyexcel/pyexcel/issues/105>`_,
   remove gease from setup_requires, introduced by 0.5.1.
#. remove python2.6 test support

0.5.1 - 20.10.2017
--------------------------------------------------------------------------------

added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. `pyexcel#103 <https://github.com/pyexcel/pyexcel/issues/103>`_, include
   LICENSE file in MANIFEST.in, meaning LICENSE file will appear in the released
   tar ball.

0.5.0 - 30.08.2017
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. put dependency on pyexcel-io 0.5.0, which uses cStringIO instead of StringIO.
   Hence, there will be performance boost in handling files in memory.

Relocated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. All ods type conversion code lives in pyexcel_io.service module

0.4.1 - 25.08.2017
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. `pyexcel#23 <https://github.com/pyexcel/pyexcel/issues/23>`_, handle
   unseekable stream given by http response
#. PR `#22 <https://github.com/pyexcel/pyexcel-ods/pull/22>`_, hanle white
   spaces in a cell.

0.4.0 - 19.06.2017
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. `pyexcel#14 <https://github.com/pyexcel/pyexcel/issues/14>`_, close file
   handle
#. pyexcel-io plugin interface now updated to use `lml
   <https://github.com/chfw/lml>`_.

0.3.3 - 07.05.2017
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. issue `pyexcel#19 <https://github.com/pyexcel/pyexcel/issues/19>`_, not all
   texts in a multi-node cell were extracted.

0.3.2 - 13.04.2017
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. issue `pyexcel#17 <https://github.com/pyexcel/pyexcel/issues/17>`_, empty new
   line is ignored
#. issue `pyexcel#6 <https://github.com/pyexcel/pyexcel/issues/6>`_,
   PT288H00M00S is valid duration

0.3.1 - 02.02.2017
--------------------------------------------------------------------------------

Added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Recognize currency type

0.3.0 - 22.12.2016
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Code refactoring with pyexcel-io v 0.3.0

0.2.2 - 24.10.2016
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. issue `pyexcel#14 <https://github.com/pyexcel/pyexcel/issues/14>`_, index
   error when reading a ods file that has non-uniform columns repeated property.

0.2.1 - 31.08.2016
--------------------------------------------------------------------------------

Added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. support pagination. two pairs: start_row, row_limit and start_column,
   column_limit help you deal with large files.
#. use odfpy 1.3.3 as compulsory package. offically support python 3

0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

Added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. By default, `float` will be converted to `int` where fits. `auto_detect_int`,
   a flag to switch off the autoatic conversion from `float` to `int`.
#. 'library=pyexcel-ods' was added so as to inform pyexcel to use it instead of
   other libraries, in the situation where multiple plugins were installed.

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. support the auto-import feature of pyexcel-io 0.2.0

0.1.1 - 30.01.2016
--------------------------------------------------------------------------------

Added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. 'streaming' is an extra option given to get_data. Only when 'streaming' is
   explicitly set to True, the data will be consisted of generators, hence will
   break your existing code.
#. uses yield in to_array and returns a generator
#. support multi-line text cell #5
#. feature migration from pyexcel-ods3 pyexcel/pyexcel-ods3#5

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. compatibility with pyexcel-io 0.1.1

0.0.12 - 10.10.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Bug fix: excessive trailing columns with empty values

0.0.11 - 26.09.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Complete fix for libreoffice datetime field

0.0.10 - 15.09.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Bug fix: date field could have datetime from libreoffice

0.0.9 - 21.08.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Bug fix: utf-8 string throw unicode exceptions

0.0.8 - 28.06.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Pin dependency odfpy 0.9.6 to avoid buggy odfpy 1.3.0

0.0.7 - 28.05.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Bug fix: "number-columns-repeated" is now respected

0.0.6 - 21.05.2015
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. get_data and save_data are seen across pyexcel-* extensions. remember them
   once and use them across all extensions.

0.0.5 - 22.02.2015
--------------------------------------------------------------------------------

Added
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. Loads only one sheet from a multiple sheet book
#. Use New BSD License

0.0.4 - 14.12.2014
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. IO interface update as pyexcel-io introduced keywords.
#. initial release

0.0.3 - 08.12.2014
--------------------------------------------------------------------------------

Updated
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#. IO interface update as pyexcel-io introduced keywords.
#. initial release
