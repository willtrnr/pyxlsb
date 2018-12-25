import os
import sys
import xml.etree.ElementTree as ElementTree
from . import recordtypes as rt
from .recordreader import RecordReader

if sys.version_info > (3,):
    xrange = range


class Cell(object):
    __slots__ = ['r', 'c', 'v', 'f']

    def __init__(self, r, c, v, f):
        self.r = r
        self.c = c
        self.v = v
        self.f = f

    def __repr__(self):
        return 'Cell(r={}, c={}, v={}, f={})'.format(self.r, self.c, self.c, self.f)


class Worksheet(object):
    def __init__(self, workbook, name, fp, rels_fp=None):
        self.workbook = workbook
        self.name = name
        self._reader = RecordReader(fp)
        self._rels_fp = rels_fp
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def __iter__(self):
        return self.rows()

    def _parse(self):
        self.dimension = None
        self.cols = list()
        self.rels = dict()
        self.hyperlinks = dict()
        self._data_offset = 0

        self._reader.seek(0, os.SEEK_SET)
        for rectype, rec in self._reader:
            if rectype == rt.WS_DIM:
                self.dimension = rec
            elif rectype == rt.COL_INFO:
                self.cols.append(rec)
            elif rectype == rt.BEGIN_SHEET_DATA:
                self._data_offset = self._reader.tell()
                if self._rels_fp is None:
                    break
            elif rectype == rt.H_LINK and self._rels_fp is not None:
                for r in xrange(rec.h):
                    for c in xrange(rec.w):
                        self.hyperlinks[rec.r + r, rec.c + c] = rec.rId

        if self._rels_fp is not None:
            self._rels_fp.seek(0, os.SEEK_SET)
            doc = ElementTree.parse(self._rels_fp).getroot()
            for el in doc:
                self.rels[el.attrib['Id']] = el.attrib['Target']

    def rows(self, sparse=True):
        row = None
        row_num = -1
        self._reader.seek(self._data_offset, os.SEEK_SET)
        for rectype, rec in self._reader:
            if rectype == rt.ROW_HDR and rec.r != row_num:
                if row is not None:
                    yield row
                while not sparse and row_num < rec.r - 1:
                    row_num += 1
                    yield [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
                row_num = rec.r
                row = [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
            elif rectype == rt.CELL_ISST:
                row[rec.c] = Cell(row_num, rec.c, self.workbook.get_shared_string(rec.v), rec.f)
            elif rectype >= rt.CELL_BLANK and rectype <= rt.FMLA_ERROR:
                row[rec.c] = Cell(row_num, rec.c, rec.v, rec.f)
            elif rectype == rt.END_SHEET_DATA:
                if row is not None:
                    yield row
                break

    def close(self):
        self._reader.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
