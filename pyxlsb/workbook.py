import sys
from . import recordtypes as rt
from .recordreader import RecordReader
from .stringtable import StringTable
from .styles import Styles
from .worksheet import Worksheet
from datetime import datetime, timedelta

if sys.version_info > (3,):
    basestring = (str, bytes)

SECONDS_IN_DAY = 24 * 60 * 60
MICROSECONDS_IN_DAY = 24 * 60 * 60 * 1000 * 1000


class Workbook(object):
    def __init__(self, pkg):
        self._pkg = pkg
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self.props = None
        self.sheets = list()
        self.stringtable = None

        with self._pkg.get_workbook_part() as f:
            for rectype, rec in RecordReader(f):
                if rectype == rt.WB_PROP:
                    self.props = rec
                elif rectype == rt.BUNDLE_SH:
                    self.sheets.append(rec.name)
                elif rectype == rt.END_BUNDLE_SHS:
                    break

        ssfp = self._pkg.get_sharedstrings_part()
        if ssfp is not None:
            self.stringtable = StringTable(ssfp)

        stylesfp = self._pkg.get_styles_part()
        if stylesfp is not None:
            self.styles = Styles(stylesfp)

    def get_sheet(self, idx_or_name, rels=False):
        if isinstance(idx_or_name, basestring):
            return self.get_sheet_by_name(idx_or_name, rels=rels)
        elif isinstance(idx_or_name, int):
            return self.get_sheet_by_index(idx_or_name - 1, rels)
        else:
            raise ValueError('string or int expected')

    def get_sheet_by_index(self, idx, rels=False):
        if idx < 0 or idx >= len(self.sheets):
            raise IndexError('sheet index out of range')

        fp = self._pkg.get_worksheet_part(idx + 1)
        rels_fp = self._pkg.get_worksheet_rels(idx + 1) if rels else None
        return Worksheet(self, self.sheets[idx], fp, rels_fp)

    def get_sheet_by_name(self, name, rels=False):
        n = name.lower()
        idx = next((n for n, s in enumerate(self.sheets) if s.lower() == n), -1) + 1
        return self.get_sheet_by_index(idx, rels=rels)

    def get_shared_string(self, idx):
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def convert_date(self, value):
        if not isinstance(value, (int, float)):
            return None

        era = datetime(1904 if self.props.date1904 else 1900, 1, 1, tzinfo=None)
        timeoffset = timedelta(microseconds=int((value % 1) * MICROSECONDS_IN_DAY))

        if int(value) == 0:
            return era + timeoffset

        if not self.props.date1904 and value >= 61:
            # According to Lotus 1-2-3, there is a Feb 29th 1900,
            # so we have to remove one extra day after that date
            dateoffset = timedelta(days=int(value) - 2)
        else:
            dateoffset = timedelta(days=int(value) - 1)

        return era + dateoffset + timeoffset

    def convert_time(self, value):
        if not isinstance(value, int) and not isinstance(value, float):
            return None
        return (datetime.min + timedelta(microseconds=int((value % 1) * MICROSECONDS_IN_DAY))).time()

    def close(self):
        self._pkg.close()
        if self.stringtable is not None:
            self.stringtable.close()
        if self.styles is not None:
            self.styles.close()
