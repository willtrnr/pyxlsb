import os
import sys
from . import records
from .recordreader import RecordReader
from .stringtable import StringTable
from .worksheet import Worksheet
from .xlsbpackage import XlsbPackage
from datetime import datetime, timedelta

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
        self.props = None
        self.sheets = list()
        self.stringtable = None

        with self._package.get_workbook_part() as f:
            reader = RecordReader(f, _debug=self._debug)
            for recid, rec in reader:
                if recid == records.WORKBOOKPR:
                    self.props = rec
                elif recid == records.SHEET:
                    self.sheets.append(rec.name)
                elif recid == records.SHEETS_END:
                    break

        try:
            self.stringtable = StringTable(self._package.get_sharedstrings_part(), _debug=self._debug)
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
        return Worksheet(self, fp, rels_fp, _debug=self._debug)

    def get_shared_string(self, idx):
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def convert_date(self, date):
        if not isinstance(date, int) and not isinstance(date, float):
            return None

        if self.props.date1904:
            era = datetime(1904, 1, 1, 0, 0, 0)
        else:
            era = datetime(1900, 1, 1, 0, 0, 0)

        if int(date) == 0:
            offset = timedelta(seconds=date * 24 * 60 * 60)
        elif date >= 61:
            # According to Lotus 1-2-3, Feb 29th 1900 is a real thing, therefore we have to remove one day after that date
            offset = timedelta(days=int(date) - 2, seconds=int((date % 1) * 24 * 60 * 60))
        else:
            # Feb 29th 1900 will show up as Mar 1st 1900 because Python won't handle that date
            offset = timedelta(days=int(date) - 1, seconds=int((date % 1) * 24 * 60 * 60))

        return era + offset

    def close(self):
        self._package.close()
