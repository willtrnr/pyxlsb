PyXLSB
======

PyXLSB is an Excel 2007-2010 Binary Workbook (XLSB / BIFF12) parser for Python.
The library is currently extremely limited, but functional enough for basic data extraction.

Usage
-----

The module exposes an `open_workbook(name)` method (similar to Xlrd and OpenPyXl) for opening XLSB files.
The Workbook object representing the file is returned.

```python
import pyxlsb
wb = pyxlsb.open_workbook('workbook.xlsb')
```

The Workbook object exposes a `get_sheet(idx)` method for retrieving a Worksheet instance.

```python
# Using the sheet index (1-based)
sheet = wb.get_sheet(1)

# Using the sheet name
sheet = wb.get_sheet('Sheet1')
```

Tip: A `sheets` property containing the sheet names is available on the Workbook instance.

The `rows()` method will hand out an iterator to read the worksheet rows.

```python
for row in sheet.rows():
  print row
```

Do note that dates will appear as `float`s.
You must use the `convert_date(date)` method from the `pyxlsb` module to turn them into `datetime` instances.

```python
print pyxlsb.convert_date(41235.45578)
```

Install
-------

Download or checkout the repository and run the `setup.py` script as usual.

```bash
python setup.py install
```
