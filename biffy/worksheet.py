import biff12
import os
import xml.etree.ElementTree as ElementTree
from reader import BIFF12Reader

class Worksheet(object):
  def __init__(self, fp, rels_fp=None, stringtable=None):
    super(Worksheet, self).__init__()
    self._reader = BIFF12Reader(fp=fp)
    self._rels_fp = rels_fp
    if not rels_fp is None:
      self._rels = ElementTree.parse(rels_fp).getroot()
    else:
      self._rels = None
    self._stringtable = stringtable
    self._data_offset = 0
    self.rels = dict()
    self.hyperlinks = dict()
    self._parse()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def _parse(self):
    if not self._rels is None:
      for el in self._rels:
        self.rels[el.attrib['Id']] = el.attrib['Target']

    for item in self._reader:
      if item[0] == biff12.SHEETDATA:
        self._data_offset = self._reader.tell()
        if self._rels is None:
          break
      elif item[0] == biff12.HYPERLINK and not self._rels is None:
        for r in xrange(item[1].r1, item[1].r2 + 1):
          for c in xrange(item[1].c1, item[1].c2 + 1):
            self.hyperlinks[r, c] = item[1].rId

  def rows(self):
    self._reader.seek(self._data_offset, os.SEEK_SET)
    row_num = 0
    row = []
    for item in self._reader:
      if item[0] == biff12.ROW and item[1].r != row_num:
        while row_num + 1 < item[1].r:
          yield row_num, []
          row_num += 1
        yield row_num, row
        row_num = item[1].r
        row = []
      elif item[0] >= biff12.BLANK and item[0] <= biff12.FORMULA_BOOLERR:
        while len(row) <= item[1].c:
          row.append(None)
        if item[0] == biff12.STRING and not self._stringtable is None:
          row[item[1].c] = self._stringtable[item[1].v]
        else:
          row[item[1].c] = item[1].v
      elif item[0] == biff12.SHEETDATA_END:
        break
    if row:
      yield row_num, row

  def close(self):
    self._reader.close()
    if not self._rels_fp is None:
      self._rels_fp.close()
