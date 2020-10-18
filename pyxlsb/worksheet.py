import os
import sys
import xml.etree.ElementTree as ElementTree
from . import biff12
from .reader import BIFF12Reader
from .writer import BIFF12Writer
from .handlers import BasicHandler, DimensionHandler, RowHandler
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

  def _write_every_row_preformat(self):
    writer = self._writer
    writer.write_id(0x0025)
    writer.write_bytes(b'\x01\x00\x02\x0e\x00\x80')
    writer.write_id(0x0895)
    writer.write_bytes(b'\x05\x00')
    writer.write_id(0x0026)
    BasicHandler.write(writer)

  def _auto_writerows(self, rows):
    # Without metadata at top up-to-date, this fails standalone
    writer = self._writer
    for i, row in enumerate(rows):
      self._write_every_row_preformat()
      RowHandler.write(writer, i)
      for v in row:
        biff_type, write_func = writer.handlers[type(v)]
        writer.write_id(biff_type)
        write_func(writer._writer, v)

  def write_table(self, table_data):
    writer = self._writer

    # Initial metadata
    writer.write_id(biff12.WORKSHEET)
    BasicHandler.write(writer)
    writer.write_id(biff12.SHEETPR)
    writer.write_bytes(b'\xc9\x04\x02\x00@\x00\x00\x00\x00\x00\x00\xff\xff\xff\xff\xff\xff\xff\xff\x00\x00\x00\x00')

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

    # Sheetviews metadata
    writer.write_id(biff12.SHEETVIEWS)
    BasicHandler.write(writer)
    writer.write_id(biff12.SHEETVIEW)
    writer.write_bytes(b'\x9c\x03\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00@\x00\x00\x00d\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00')
    writer.write_id(biff12.SELECTION)
    writer.write_bytes(b'\x03\x00\x00\x00\x17\x00\x00\x00\x02\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00\x17\x00\x00\x00\x17\x00\x00\x00\x02\x00\x00\x00\x02\x00\x00\x00')
    writer.write_id(biff12.SHEETVIEW_END)
    BasicHandler.write(writer)
    writer.write_id(biff12.SHEETVIEWS_END)
    BasicHandler.write(writer)

    # Remaining metadata
    self._write_every_row_preformat()
    writer.write_id(biff12.SHEETFORMATPR)
    writer.write_bytes(b'\xff\xff\xff\xff\x08\x00,\x01\x00\x00\x00\x00')

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
      for i, row in enumerate(table_data.iterrows()):  # TODO: better/faster than iterrows?
        RowHandler.write(writer, i)
        for cell, (biff_type, write_func) in zip(row, writer_handlers):
          self._write_every_row_preformat()
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
