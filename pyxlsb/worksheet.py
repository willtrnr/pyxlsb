import os
import sys
import xml.etree.ElementTree as ET
from . import recordtypes as rt
from .row import Row
from .recordreader import RecordReader

if sys.version_info > (3,):
    xrange = range


class Worksheet(object):
    """A worksheet.

    Attributes:
        workbook (Workbook): The containing workbook.
        name (str): The name of this worksheet.
    """

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
        self.cols = []
        self.rels = {}
        self.hyperlinks = {}
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
            doc = ET.parse(self._rels_fp)
            for el in doc.getroot():
                self.rels[el.attrib['Id']] = el.attrib['Target']

    def rows(self, sparse=True):
        """Get a row iterator to iterate through the cell data.

        Args:
            sparse (bool, optional): If empty rows should be skipped, defaults to True

        Yields:
            Row: The rows in this worksheet.
        """
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
                row._add_cell(rec.c, self.workbook.get_shared_string(rec.v), rec.f, rec.style)
            elif rectype >= rt.CELL_BLANK and rectype <= rt.FMLA_ERROR:
                row._add_cell(rec.c, rec.v, rec.f, rec.style)
            elif rectype == rt.END_SHEET_DATA:
                if row is not None:
                    yield row
                break

    def close(self):
        """Release the resources associated with this worksheet."""
        self._fp.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
