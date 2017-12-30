pyxlsb
======

[![PyPI](https://img.shields.io/pypi/v/pyxlsb.svg)](https://pypi.python.org/pypi/pyxlsb)
[![Build Status](https://travis-ci.org/wwwiiilll/pyxlsb.svg?branch=master)](https://travis-ci.org/wwwiiilll/pyxlsb)

`pyxlsb` is an Excel 2007-2010 Binary Workbook (xlsb) parser for Python.
The library is currently limited, but should be functional enough for basic
data extraction.

Install
-------

```bash
pip install pyxlsb
```

Usage
-----

The module exposes an `open_workbook(name)` method (similar to Xlrd and
OpenPyXl) for opening XLSB files. The Workbook object representing the file is
returned.

```python
from pyxlsb import open_workbook
with open_workbook('Book1.xlsb') as wb:
    # Do stuff with wb
```

The Workbook object exposes a `get_sheet(idx)` method for retrieving a
Worksheet instance.

```python
# Using the sheet index (1-based to match VBA)
with wb.get_sheet(1) as sheet:
    # Do stuff with sheet

# Using the sheet name
with wb.get_sheet('Sheet1') as sheet:
    # Do stuff with sheet
```

Tip: A `sheets` property containing the sheet names is available on the
Workbook instance.

The `rows()` method will hand out an iterator to read the worksheet rows. The
Worksheet object is also directly iterable and is equivalent to calling
`rows()`.

```python
# You can use .rows(sparse=True) to skip empty rows and iterate faster
for row in sheet.rows():
  print(row)
# [Cell(r=0, c=0, v='TEXT'), Cell(r=0, c=1, v=42.1337)]
```

*NOTE*: Iterating the same Worksheet instance multiple times in parallel (nested
`for`s for instance) will yield unexpected results, retrieve more instances
using `get_sheet` above.

Do note that dates will appear as floats. You must use the `convert_date(date)`
method from the corresponding Workbook instance to turn them into `datetime`.

```python
print(wb.convert_date(41235.45578))
# datetime.datetime(2012, 11, 22, 10, 56, 19)
```

*NOTE*: Using the `convert_date` in the `pyxlsb` module still works, but is
deprecated and will be removed.

Limitations
-----------

Non exhaustive list of things that are currently not supported:

  - Formulas
    - Parsing *WIP*
    - Evaluation
  - Formatting *WIP*
  - Rich text cells (formatting is lost currently)
  - Encrypted (password protected) workbooks
  - Comments and other annotations
  - ~~1904 date system~~ *Done*
  - Writing (*very* far goal)

Feel free to open issues or, even better, send PRs for these things and anything
else I might have missed, I'll try to prioritize what's most requested.
