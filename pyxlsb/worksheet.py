import os
import sys
import xml.etree.ElementTree as ElementTree
from . import biff12
from .reader import BIFF12Reader
from .writer import BIFF12Writer
from .handlers import BasicHandler, DimensionHandler
from collections import namedtuple

if sys.version_info > (3,):
  xrange = range

supported_modes = {'r', 'w'}

Cell = namedtuple('Cell', ['r', 'c', 'v'])

class Worksheet(object):
  def __init__(self, fp, rels_fp=None, stringtable=None, debug=False, mode='r'):
    super(Worksheet, self).__init__()
    self._mode = mode
    if self._mode == 'w':
      self._writer = BIFF12Writer(fp=fp, debug=debug)
    else:
      self._reader = BIFF12Reader(fp=fp, debug=debug)
    self._rels_fp = rels_fp
    if not rels_fp is None:
      self._rels = ElementTree.parse(rels_fp).getroot()
    else:
      self._rels = None
    self._stringtable = stringtable
    self._data_offset = 0
    self.dimension = None
    self.cols = list()
    self.rels = dict()
    self.hyperlinks = dict()
    if self._mode == 'r':
      self._parse()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def __iter__(self):
    if self._mode != 'r':
      raise NotImplementedError('Not opened for reading')
    return self.rows()

  def _parse(self):
    assert self._mode == 'r', 'Not opened for reading'
    if self._rels is not None:
      for el in self._rels:
        self.rels[el.attrib['Id']] = el.attrib['Target']

    for item in self._reader:
      if item[0] == biff12.DIMENSION:
        self.dimension = item[1]
      elif item[0] == biff12.COL:
        self.cols.append(item[1])
      elif item[0] == biff12.SHEETDATA:
        self._data_offset = self._reader.tell()
        if self._rels is None:
          break
      elif item[0] == biff12.HYPERLINK and self._rels is not None:
        for r in xrange(item[1].h):
          for c in xrange(item[1].w):
            self.hyperlinks[item[1].r + r, item[1].c + c] = item[1].rId

  def rows(self, sparse=False):
    assert self._mode == 'r', 'Not opened for reading'
    self._reader.seek(self._data_offset, os.SEEK_SET)
    row_num = -1
    row = None
    for item in self._reader:
      if item[0] == biff12.ROW and item[1].r != row_num:
        if row is not None:
          yield row
        if not sparse:
          while row_num < item[1].r - 1:
            row_num += 1
            yield [Cell(row_num, i, None) for i in xrange(self.dimension.c + self.dimension.w)]
        row_num = item[1].r
        row = [Cell(row_num, i, None) for i in xrange(self.dimension.c + self.dimension.w)]
      elif item[0] >= biff12.BLANK and item[0] <= biff12.FORMULA_BOOLERR:
        if item[0] == biff12.STRING and self._stringtable is not None:
          row[item[1].c] = Cell(row_num, item[1].c, self._stringtable[item[1].v])
        else:
          row[item[1].c] = Cell(row_num, item[1].c, item[1].v)
      elif item[0] == biff12.SHEETDATA_END:
        if row is not None:
          yield row
        break

  def _auto_writerows(self, rows):
    # Without metadata at top up-to-date, this fails standalone
    writer = self._writer
    for row in rows:
      writer.write_id(biff12.ROW)
      for v in row:
        biff_type, write_func = writer.handlers[type(v)]
        writer.write_id(biff_type)
        write_func(writer._writer, v)

  def write_table(self, table_data):
    writer = self._writer

    # Dimension metadata
    try:
      # Handle table_data as something like a pandas.DataFrame or numpy.ndarray
      num_rows, num_cols = table_data.shape
    except Exception:
      num_rows = len(table_data)
      num_cols = 0
      for row in table_data:
        num_cols = max(num_cols, len(row))
    writer.write_id(biff12.DIMENSION)
    DimensionHandler.write(writer, num_rows, num_cols)

    # Sheet data
    writer.write_id(biff12.SHEETDATA)
    BasicHandler.write(writer)
    try:
      # Handle table_data as something like a pandas.DataFrame or numpy.ndarray
      try:
        # Handle table_data as something like a pandas.DataFrame
        writer_handlers = [writer.handlers[dt] for dt in table_data.dtypes]
      except AttributeError:
        # Handle table_data as something like a numpy.ndarray
        writer_handlers = [writer.handlers[table_data.dtype]] * num_cols
      for row in table_data.iterrows():  # TODO: better/faster than iterrows?
        for cell, (biff_type, write_func) in zip(row, writer_handlers):
          writer.write_id(biff_type)
          write_func(writer._writer, cell)
    except AttributeError:
      # Handle table_data as a list of rows (each themselves a list)
      self._auto_writerows(table_data)

    # Ending metadata
    writer.write_id(biff12.SHEETDATA_END)
    BasicHandler.write(writer)
    writer.write_id(biff12.WORKSHEET_END)
    BasicHandler.write(writer)

  def close(self):
    try:
      self._reader.close()
    except AttributeError:
      self._writer.close()
    if self._rels_fp is not None:
      self._rels_fp.close()
