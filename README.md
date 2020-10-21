pyxlsb
======

[![PyPI](https://img.shields.io/pypi/v/pyxlsb.svg)](https://pypi.python.org/pypi/pyxlsb)

`pyxlsb` is an Excel 2007-2010 Binary Workbook (xlsb) parser for Python.
The library is currently extremely limited, but functional enough for basic
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
# Using the sheet index (1-based)
with wb.get_sheet(1) as sheet:
    # Do stuff with sheet

# Using the sheet name
with wb.get_sheet('Sheet1') as sheet:
    # Do stuff with sheet
```

Tip: A `sheets` property containing the sheet names is available on the
Workbook instance.

The `rows()` method will hand out an iterator to read the worksheet rows.

```python
# You can use .rows(sparse=True) to skip empty rows
for row in sheet.rows():
  print(row)
# [Cell(r=0, c=0, v='TEXT'), Cell(r=0, c=1, v=42.1337)]
```

Do note that dates will appear as floats. You must use the `convert_date(date)`
method from the `pyxlsb` module to turn them into `datetime` instances.

```python
from pyxlsb import convert_date
print(convert_date(41235.45578))
# datetime.datetime(2012, 11, 22, 10, 56, 19)
```

Basic support for writing XLSB files is available, enabling the creation of new sheets populated with data from `pandas.DataFrame`, `numpy.ndarray`, or `list`-of-`list` objects.  (The `pandas` and `numpy` packages are purely optional dependencies of `pyxlsb`.)

```python
from pyxlsb import open_workbook
import pandas as pd
with open_workbook('Book2.xlsb', 'w') as wb:
    with wb.create_sheet('Sheet1') as ws:
        df = pd.DataFrame(
            [
                ['blue', 4, -273.154],
                ['azul', 5, -195.79],
                ['blau', 6, -78.5],
            ],
            columns=['colors', 'points', 'temp']
        )
        ws.write_table(df)
```

```python
from pyxlsb import open_workbook
import numpy as np
wb = open_workbook('Book3.xlsb', 'w')
ws = wb.create_sheet('Sheet1')
try:
    a = np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9.9]])
    ws.write_table(a)
finally:
    ws.close()  # Ensure sheets are closed before writing any other sheets.
ws = wb.create_sheet('Sheet2')
try:
    data = [
        ['blue', 4, -273.154],
        ['azul', 5, -195.79],
        ['blau', 6, -78.5],
    ]
    ws.write_table(data)
finally:
    ws.close()
wb.close()  # Ensure all sheets are closed before closing the workbook.
```