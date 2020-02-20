import os
import sys
import xml.etree.ElementTree as ElementTree
from . import recordtypes as rt
from .row import Row
from .recordreader import RecordReader
from .conv import convert_value

if sys.version_info > (3,):
    xrange = range


class Worksheet(object):
    def __init__(self, workbook, name, fp, rels_fp=None):
        self.workbook = workbook
        self.name = name
        self._fp = fp
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

        self._fp.seek(0, os.SEEK_SET)
        for rectype, rec in RecordReader(self._fp):
            if rectype == rt.WS_DIM:
                self.dimension = rec
            elif rectype == rt.COL_INFO:
                self.cols.append(rec)
            elif rectype == rt.BEGIN_SHEET_DATA:
                self._data_offset = self._fp.tell()
                if self._rels_fp is None:
                    break
            elif rectype == rt.H_LINK and self._rels_fp is not None:
                for r in xrange(rec.h):
                    for c in xrange(rec.w):
                        self.hyperlinks[rec.r + r, rec.c + c] = rec.rId

        if self._rels_fp is not None:
            self._rels_fp.seek(0, os.SEEK_SET)
            doc = ElementTree.parse(self._rels_fp)
            for el in doc.getroot():
                self.rels[el.attrib['Id']] = el.attrib['Target']

    def rows(self, sparse=True):
        row = None
        row_num = -1
        self._fp.seek(self._data_offset, os.SEEK_SET)
        for rectype, rec in RecordReader(self._fp):
            if rectype == rt.ROW_HDR and rec.r != row_num:
                if row is not None:
                    yield row
                while not sparse and row_num < rec.r - 1:
                    row_num += 1
                    yield Row(self, row_num)
                row_num = rec.r
                row = Row(self, row_num)
            elif rectype == rt.CELL_ISST:
                row._add_cell(rec.c, self.workbook.get_shared_string(rec.v), self.workbook.get_shared_string(rec.v), rec.f, rec.style, self.workbook.styles.get_style(rec.style))
            elif rectype >= rt.CELL_BLANK and rectype <= rt.FMLA_ERROR:
                style_fmt = self.workbook.styles.get_style(rec.style)
                row._add_cell(rec.c, rec.v, convert_value(rec.v, style_fmt.dtype), rec.f, rec.style, style_fmt)
            elif rectype == rt.END_SHEET_DATA:
                if row is not None:
                    yield row
                break

    def close(self):
        self._fp.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
