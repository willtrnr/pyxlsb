import os
import sys
import xml.etree.ElementTree as ElementTree
from . import records
from .record_reader import RecordReader
from collections import namedtuple

if sys.version_info > (3,):
    xrange = range

Cell = namedtuple('Cell', ['r', 'c', 'v', 'f'])

class Worksheet(object):
    def __init__(self, fp, rels_fp=None, stringtable=None, _debug=False):
        super(Worksheet, self).__init__()
        self._reader = RecordReader(fp=fp, _debug=_debug)
        self._rels_fp = rels_fp
        self._rels = ElementTree.parse(rels_fp).getroot() if rels_fp else None
        self._stringtable = stringtable
        self._data_offset = 0

        self.dimension = None
        self.cols = list()
        self.rels = dict()
        self.hyperlinks = dict()

        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def __iter__(self):
        return self.rows()

    def _parse(self):
        if self._rels is not None:
            for el in self._rels:
                self.rels[el.attrib['Id']] = el.attrib['Target']

        for recid, rec in self._reader:
            if recid == records.DIMENSION:
                self.dimension = rec
            elif recid == records.COL:
                self.cols.append(rec)
            elif recid == records.SHEETDATA:
                self._data_offset = self._reader.tell()
                if self._rels is None:
                    break
            elif recid == records.HYPERLINK and self._rels is not None:
                for r in xrange(rec.h):
                    for c in xrange(rec.w):
                        self.hyperlinks[rec.r + r, rec.c + c] = rec.rId

    def rows(self, sparse=False):
        self._reader.seek(self._data_offset, os.SEEK_SET)
        row_num = -1
        row = None
        for recid, rec in self._reader:
            if recid == records.ROW and rec.r != row_num:
                if row is not None:
                    yield row
                if not sparse:
                    while row_num < rec.r - 1:
                        row_num += 1
                        yield [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
                row_num = rec.r
                row = [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
            elif recid >= records.BLANK and recid <= records.FORMULA_BOOLERR:
                if recid == records.STRING and self._stringtable is not None:
                    row[rec.c] = Cell(row_num, rec.c, self._stringtable[rec.v], rec.f)
                else:
                    row[rec.c] = Cell(row_num, rec.c, rec.v, rec.f)
            elif recid == records.SHEETDATA_END:
                if row is not None:
                    yield row
                break

    def close(self):
        self._reader.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
