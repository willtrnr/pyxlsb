import sys
import logging
import warnings
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from . import recordtypes as rt
from .recordreader import RecordReader
from .stringtable import StringTable
from .styles import Styles
from .worksheet import Worksheet

if sys.version_info > (3,):
    basestring = (str, bytes)
    long = int

_MICROSECONDS_IN_DAY = 24 * 60 * 60 * 1000 * 1000


class Workbook(object):
    """The main Workbook class providing worksheets and other metadata.

    Args:
        pkg (WorkbookPackage): The package driver backing this workbook.
    """

    def __init__(self, pkg):
        self._pkg = pkg
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, typ, value, traceback):
        self.close()

    def _parse(self):
        self.props = None
        self.stringtable = None
        self.styles = None
        self._sheets = []

        rels = {}
        with self._pkg.get_workbook_rels() as f:
            for el in ET.parse(f).getroot():
                rels[el.attrib['Id']] = el.attrib['Target']

        with self._pkg.get_workbook_part() as f:
            for rectype, rec in RecordReader(f):
                if rectype == rt.WB_PROP:
                    self.props = rec
                elif rectype == rt.BUNDLE_SH:
                    self._sheets.append((rec.name, rels[rec.rId]))
                elif rectype == rt.END_BUNDLE_SHS:
                    break

        ssfp = self._pkg.get_sharedstrings_part()
        if ssfp is not None:
            self.stringtable = StringTable(ssfp)

        stylesfp = self._pkg.get_styles_part()
        if stylesfp is not None:
            self.styles = Styles(stylesfp)

    @property
    def sheets(self):
        """:obj:`list` of :obj:`str`: List of sheet names in this workbook."""
        return [v[0] for v in self._sheets]

    def get_sheet(self, idx_or_name, with_rels=False):
        """Get a sheet by name or 1-based index.

        Args:
            idx_or_name (:obj:`int` or :obj:`str`): The 1-based index or name of the sheet
            with_rels (:obj:`bool`, optional): If the relationships should be parsed, defaults to False

        Returns:
            Worksheet: The corresponding worksheet instance.

        Raises:
            IndexError: When the index or name is invalid.
            ValueError: For an incompatible index or name type.

        .. deprecated:: 1.1.0
            Will be removed in 1.2.0. Use either ``get_sheet_by_index`` or ``get_sheet_by_name``
            instead.
        """
        warnings.warn("get_sheet was replaced with get_sheet_by_name/index", DeprecationWarning)
        if isinstance(idx_or_name, basestring):
            return self.get_sheet_by_name(idx_or_name, with_rels)
        elif isinstance(idx_or_name, (int, long)):
            return self.get_sheet_by_index(idx_or_name - 1, with_rels)
        else:
            raise ValueError('string or int expected')

    def get_sheet_by_index(self, idx, with_rels=False):
        """Get a worksheet by index.

        Args:
            idx (int): The index of the sheet.
            with_rels (:obj:`bool`, optional): If the relationships should be parsed, defaults to False

        Returns:
            Worksheet: The corresponding worksheet instance.

        Raises:
            IndexError: When the provided index is out of range.
        """
        if idx < 0 or idx >= len(self._sheets):
            raise IndexError('sheet index out of range')

        target = self._sheets[idx][1]

        fp = self._pkg.get_file('xl/' + target)
        if with_rels:
            parts = target.split('/')
            rels_fp = self._pkg.get_file('xl/' + '/'.join(parts[:-1] + ['_rels'] + parts[-1:]) + '.rels')
        else:
            rels_fp = None
        return Worksheet(self, self._sheets[idx][0], fp, rels_fp)

    def get_sheet_by_name(self, name, with_rels=False):
        """Get a worksheet by its name.

        Args:
            name (str): The name of the sheet.
            with_rels (:obj:`bool`, optional): If the relationships should be parsed, defaults to False

        Returns:
            Worksheet: The corresponding worksheet instance.

        Raises:
            IndexError: When the provided name is invalid.
        """
        n = name.lower()
        idx = next((i for i, (s, _) in enumerate(self._sheets) if s.lower() == n), -1)
        return self.get_sheet_by_index(idx, with_rels=with_rels)

    def get_shared_string(self, idx):
        """Get a string in the shared string table.

        Args:
            idx (int): The index of the string in the shared string table.

        Returns:
            str: The string at the index in table table or None if not found.
        """
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def convert_date(self, value):
        """Convert an Excel numeric date value to a ``datetime`` instance.

        Args:
            value (:obj:`int` or :obj:`float`): The Excel date value.

        Returns:
            datetime.datetime: The equivalent datetime instance or None when invalid.
        """
        if not isinstance(value, (int, long, float)):
            return None

        era = datetime(1904 if self.props.date1904 else 1900, 1, 1, tzinfo=None)
        timeoffset = timedelta(microseconds=long((value % 1) * _MICROSECONDS_IN_DAY))

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
        """Convert an Excel date fraction time value to a ``time`` instance.

        Args:
            value (float): The Excel fractional day value.

        Returns:
            datetime.time: The equivalent time instance or None if invalid.
        """
        if not isinstance(value, (int, long, float)):
            return None
        return (datetime.min + timedelta(microseconds=long((value % 1) * _MICROSECONDS_IN_DAY))).time()

    def close(self):
        """Release the resources associated with this workbook."""
        self._pkg.close()
        if self.stringtable is not None:
            self.stringtable.close()
        if self.styles is not None:
            self.styles.close()
