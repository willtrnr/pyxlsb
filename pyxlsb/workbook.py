import os
import sys
from . import biff12
from .reader import BIFF12Reader
from .writer import BIFF12Writer
from .stringtable import StringTable
from .worksheet import Worksheet, supported_modes
from tempfile import TemporaryFile

if sys.version_info > (3,):
  basestring = (str, bytes)

class Workbook(object):
  def __init__(self, fp, debug=False):
    super(Workbook, self).__init__()
    self._zf = fp
    self._mode = self._zf.mode
    assert self._mode in supported_modes
    self._debug = debug
    self.sheets = list()
    self.stringtable = None
    if self._mode == 'r':
      self._parse()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def _parse(self):
    assert self._mode == 'r', 'Not opened for reading'
    with TemporaryFile() as temp:
      with self._zf.open('xl/workbook.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      reader = BIFF12Reader(fp=temp, debug=self._debug)
      for item in reader:
        if item[0] == biff12.SHEET:
          self.sheets.append(item[1].name)
        elif item[0] == biff12.SHEETS_END:
          break

    try:
      temp = TemporaryFile()
      with self._zf.open('xl/sharedStrings.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      self.stringtable = StringTable(fp=temp)
    except KeyError:
      pass

  def get_sheet(self, idx, rels=False):
    assert self._mode == 'r', 'Not opened for reading'
    if isinstance(idx, basestring):
      idx = [s.lower() for s in self.sheets].index(idx.lower()) + 1
    if idx < 1 or idx > len(self.sheets):  # TODO: this check seems unnecessary given behavior of list.index
      raise IndexError('sheet index out of range')

    temp = TemporaryFile()
    with self._zf.open('xl/worksheets/sheet{}.bin'.format(idx), 'r') as zf:
      temp.write(zf.read())
      temp.seek(0, os.SEEK_SET)

    if rels:
      rels_temp = TemporaryFile()
      with self._zf.open('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx), 'r') as zf:
        rels_temp.write(zf.read())
        rels_temp.seek(0, os.SEEK_SET)
    else:
      rels_temp = None

    return Worksheet(fp=temp, rels_fp=rels_temp, stringtable=self.stringtable, debug=self._debug)

  def create_sheet(self, sheet_name):
    assert self._mode == 'w', 'Not opened for writing'
    assert sheet_name not in self.sheets, 'Duplicate sheet name'

    # TODO: update (not only create) xl/workbook.bin
    self.sheets.append(sheet_name)
    with self._zf.open('xl/workbook.bin', 'w') as workbook_zf:
      writer = BIFF12Writer(fp=workbook_zf, debug=self._debug)
      for sn in self.sheets:
        writer.write_id(biff12.SHEET)
        writer.write_string(sn)
      writer.write_id(biff12.SHEETS_END)

    sheet_idx = len(self.sheets)
    sheet_zf = self._zf.open('xl/worksheets/sheet{}.bin'.format(sheet_idx), 'w')
    try:
      sheet = Worksheet(fp=sheet_zf, rels_fp=None, stringtable=self.stringtable, debug=self._debug, mode='w')
    except BaseException:
      sheet_zf.close()
      raise

    return sheet

  def close(self):
    self._zf.close()
