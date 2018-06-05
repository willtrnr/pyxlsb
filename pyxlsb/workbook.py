import sys
from . import records
from .recordreader import RecordReader
from .stringtable import StringTable
from .styles import Styles
from .worksheet import Worksheet
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

        try:
            self.styles = Styles(self._package.get_styles_part(), _debug=self._debug)
        except KeyError:
            self.styles = None

    def get_sheet(self, idx, rels=False):
        if isinstance(idx, basestring):
            idx = idx.lower()
            idx = next((n for n, s in enumerate(self.sheets) if s.lower() == idx), -1) + 1
        if idx < 0 or idx > len(self.sheets):
            raise IndexError('sheet index out of range')

        fp = self._package.get_worksheet_part(idx)
        rels_fp = self._package.get_worksheet_rels(idx) if rels else None
        return Worksheet(self, fp, rels_fp, _debug=self._debug)

    def get_shared_string(self, idx):
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def convert_date(self, value):
        if not isinstance(value, int) and not isinstance(value, float):
            return None
        era = datetime(1904 if self.props.date1904 else 1900, 1, 1, tzinfo=None)
        timeoffset = timedelta(seconds=int((value % 1) * 24 * 60 * 60))
        if int(value) == 0:
            return era + timeoffset
        elif not self.props.date1904 and int(value) >= 61:
            # According to Lotus 1-2-3, there is a Feb 29th 1900,
            # so we have to remove one day after that date
            dateoffset = timedelta(days=int(value) - 2)
        else:
            dateoffset = timedelta(days=int(value) - 1)
        return era + dateoffset + timeoffset

    def convert_time(self, value):
        if not isinstance(value, int) and not isinstance(value, float):
            return None
        return (datetime.min + timedelta(seconds=int((value % 1) * 24 * 60 * 60))).time()

    def close(self):
        self._package.close()
        if self.stringtable is not None:
            self.stringtable.close()
        if self.styles is not None:
            self.styles.close()
