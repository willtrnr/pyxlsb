import os
import sys
import xml.etree.ElementTree as ElementTree
from . import records
from .recordreader import RecordReader
from collections import namedtuple

if sys.version_info > (3,):
    xrange = range

# TODO: Make a real class of this?
Cell = namedtuple('Cell', ['r', 'c', 'v', 'f'])


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
            if rectype == records.WS_DIM:
                self.dimension = rec
            elif rectype == records.COL_INFO:
                self.cols.append(rec)
            elif rectype == records.BEGIN_SHEET_DATA:
                self._data_offset = self._reader.tell()
                if self._rels_fp is None:
                    break
            elif rectype == records.H_LINK and self._rels_fp is not None:
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
            if rectype == records.ROW_HDR and rec.r != row_num:
                if row is not None:
                    yield row
                while not sparse and row_num < rec.r - 1:
                    row_num += 1
                    yield [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
                row_num = rec.r
                row = [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
            elif rectype == records.CELL_ISST:
                row[rec.c] = Cell(row_num, rec.c, self.workbook.get_shared_string(rec.v), rec.f)
            elif rectype >= records.CELL_BLANK and rectype <= records.FMLA_ERROR:
                row[rec.c] = Cell(row_num, rec.c, rec.v, rec.f)
            elif rectype == records.END_SHEET_DATA:
                if row is not None:
                    yield row
                break

    def close(self):
        self._reader.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
