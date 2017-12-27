import os
import sys
from . import records
from .recordreader import RecordReader
from .stringtable import StringTable
from .worksheet import Worksheet
from .xlsbpackage import XlsbPackage

if sys.version_info > (3,):
    basestring = (str, bytes)

class Workbook(object):
    def __init__(self, package, _debug=False):
        super(Workbook, self).__init__()
        self._package = package
        self._debug = _debug

        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self.sheets = list()
        self.stringtable = None

        with self._package.get_workbook_part() as f:
            reader = RecordReader(f, _debug=self._debug)
            for recid, rec in reader:
                if recid == records.SHEET:
                    self.sheets.append(rec.name)
                elif recid == records.SHEETS_END:
                    break

        try:
            self.stringtable = StringTable(fp=self._package.get_sharedstrings_part(), _debug=self._debug)
        except KeyError:
            self.stringtable = None

    def get_sheet(self, idx, rels=False):
        if isinstance(idx, basestring):
            idx = idx.lower()
            idx = next((n for n, s in enumerate(self.sheets) if s.lower() == idx), -1) + 1
        if idx < 1 or idx > len(self.sheets):
            raise IndexError('sheet index out of range')

        fp = self._package.get_worksheet_part(idx)
        rels_fp = self._package.get_worksheet_rels(idx) if rels else None
        return Worksheet(self, fp, rels_fp=rels_fp, _debug=self._debug)

    def get_shared_string(self, idx):
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def close(self):
        self._zf.close()
