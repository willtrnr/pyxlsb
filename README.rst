pyxlsb
======

|PyPI| |Build Status|

``pyxlsb`` is an Excel 2007-2010 Binary Workbook (xlsb) parser for Python. The
library is currently very limited, but should be functional enough for basic
data extraction or conversion.

Install
-------

.. code:: bash

   pip install pyxlsb

Usage
-----

Also see the `API reference`_.

The module exposes an ``open_workbook(name)`` method (similar to Xlrd and
OpenPyXl) for opening XLSB files. The Workbook object representing the file is
returned.

.. code:: python

   from pyxlsb import open_workbook
   with open_workbook('Book1.xlsb') as wb:
       # Do stuff with wb

The Workbook object exposes a ``get_sheet_by_index(idx)`` and
``get_sheet_by_name(name)`` method to retrieve Worksheet instances.

.. code:: python

   # Using the sheet index (0-based, unlike VBA)
   with wb.get_sheet_by_index(0) as sheet:
       # Do stuff with sheet

   # Using the sheet name
   with wb.get_sheet_by_name('Sheet1') as sheet:
       # Do stuff with sheet

*NOTE*: The 1-based ``get_sheet(idx_or_name)`` method still works, but is
deprecated.

A ``sheets`` property containing the sheet names is available on the Workbook
instance.

The ``rows()`` method will hand out an iterator to read the worksheet rows. The
Worksheet object is also directly iterable and is equivalent to calling
``rows()``.

.. code:: python

   # You can use .rows(sparse=False) to include empty rows
   for row in sheet.rows():
       print(row)
   # [Cell(r=0, c=0, v='TEXT'), Cell(r=0, c=1, v=42.1337)]

*NOTE*: Iterating the same Worksheet instance multiple times in parallel (nested
``for`` for instance) will yield unexpected results, retrieve more instances
instead.

Note that dates will appear as floats. You must use the ``convert_date(date)``
method from the corresponding Workbook instance to turn them into ``datetime``.

.. code:: python

   print(wb.convert_date(41235.45578))
   # datetime.datetime(2012, 11, 22, 10, 56, 19)

*NOTE*: Using the ``convert_date`` in the ``pyxlsb`` module still works, but is
deprecated and will be removed.

Example
-------

Converting a workbook to CSV:

.. code:: python

   import csv
   from pyxlsb import open_workbook

   with open_workbook('Book1.xlsb') as wb:
       for name in wb.sheets:
           with wb.get_sheet_by_name(name) as sheet:
               with open(name + '.csv', 'w') as f:
                   writer = csv.writer(f)
                   for row in sheet.rows():
                       writer.writerow([c.v for c in row])

Limitations
-----------

Non exhaustive list of things that are currently not supported:

-  Formulas

   -  Parsing *WIP*
   -  Evaluation

-  Style and formatting *WIP*
-  Rich text cells (formatting is lost, but getting the text works)
-  Encrypted (password protected) workbooks
-  Comments and other annotations
-  Writing (*very* far goal)

Feel free to open issues or, even better, submit PRs for these things and
anything else I might have missed, I'll try to prioritize what's most requested.

.. |PyPI| image:: https://img.shields.io/pypi/v/pyxlsb.svg
   :target: https://pypi.python.org/pypi/pyxlsb
.. |Build Status| image:: https://travis-ci.org/willtrnr/pyxlsb.svg?branch=master
   :target: https://travis-ci.org/willtrnr/pyxlsb
.. _API reference: https://willtrnr.github.io/pyxlsb/
