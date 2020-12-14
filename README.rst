pyxlsb
======

|PyPI|

``pyxlsb`` is an Excel 2007-2010 Binary Workbook (xlsb) parser for
Python. The library is currently extremely limited, but functional
enough for basic data extraction.

Install
-------

.. code:: sh

   pip install pyxlsb

Usage
-----

The module exposes an ``open_workbook(name)`` method (similar to Xlrd
and OpenPyXl) for opening XLSB files. The Workbook object representing
the file is returned.

.. code:: python

   from pyxlsb import open_workbook
   with open_workbook('Book1.xlsb') as wb:
       # Do stuff with wb

The Workbook object exposes a ``get_sheet(idx)`` method for retrieving a
Worksheet instance.

.. code:: python

   # Using the sheet index (1-based)
   with wb.get_sheet(1) as sheet:
       # Do stuff with sheet

   # Using the sheet name
   with wb.get_sheet('Sheet1') as sheet:
       # Do stuff with sheet

Tip: A ``sheets`` property containing the sheet names is available on
the Workbook instance.

The ``rows()`` method will hand out an iterator to read the worksheet
rows.

.. code:: python

   # You can use .rows(sparse=True) to skip empty rows
   for row in sheet.rows():
       print(row)
   # [Cell(r=0, c=0, v='TEXT'), Cell(r=0, c=1, v=42.1337)]

Do note that dates will appear as floats. You must use the
``convert_date(date)`` method from the ``pyxlsb`` module to turn them
into ``datetime`` instances.

.. code:: python

   from pyxlsb import convert_date
   print(convert_date(41235.45578))
   # datetime.datetime(2012, 11, 22, 10, 56, 19)

.. |PyPI| image:: https://img.shields.io/pypi/v/pyxlsb.svg
   :target: https://pypi.python.org/pypi/pyxlsb
