import os
import sys
from . import records
from .record_reader import RecordReader
from .stringtable import StringTable
from .worksheet import Worksheet
from tempfile import TemporaryFile

if sys.version_info > (3,):
    basestring = (str, bytes)

class Workbook(object):
    def __init__(self, fp, _debug=False):
        super(Workbook, self).__init__()
        self._zf = fp
        self._debug = _debug

        self.sheets = list()
        self.stringtable = None

        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        with TemporaryFile() as temp:
            with self._zf.open('xl/workbook.bin', 'r') as zf:
                temp.write(zf.read())
                temp.seek(0, os.SEEK_SET)
            reader = RecordReader(fp=temp, _debug=self._debug)
            for recid, rec in reader:
                if recid == records.SHEET:
                    self.sheets.append(rec.name)
                elif recid == records.SHEETS_END:
                    break

        try:
            temp = TemporaryFile()
            with self._zf.open('xl/sharedStrings.bin', 'r') as zf:
                temp.write(zf.read())
                temp.seek(0, os.SEEK_SET)
            self.stringtable = StringTable(fp=temp, _debug=self._debug)
        except KeyError:
            pass

    def get_sheet(self, idx, rels=False):
        if isinstance(idx, basestring):
            idx = [s.lower() for s in self.sheets].index(idx.lower()) + 1
        if idx < 1 or idx > len(self.sheets):
            raise IndexError('sheet index out of range')

        temp = TemporaryFile()
        with self._zf.open('xl/worksheets/sheet{}.bin'.format(idx), 'r') as zf:
            temp.write(zf.read())
            temp.seek(0, os.SEEK_SET)

        if rels:
            rels_temp = TemporaryFile()
            with self._zf.open('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx), 'r') as zf:
                rels_temp.write(zf.read())
                rels_temp.seek(0, os.SEEK_SET)
        else:
            rels_temp = None

        return Worksheet(fp=temp, rels_fp=rels_temp, stringtable=self.stringtable, _debug=self._debug)

    def close(self):
        self._zf.close()
