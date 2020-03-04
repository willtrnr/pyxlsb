import os
from . import recordtypes as rt
from .recordreader import RecordReader
from .records import FormatRecord


class Format(object):
    __slots__ = ('code', 'is_builtin', 'is_date_format')

    def __init__(self, code, is_builtin=False):
        self.code = code
        self.is_builtin = is_builtin

        self.is_date_format = False
        if code is not None:
            for c in ('y', 'm', 'd', 'h', 's', 'AM', 'PM'):
                if c in code:
                    self.is_date_format = True
                    break


    def __repr__(self):
        return 'Format(code={}, is_builtin={})' \
            .format(self.code, self.is_builtin)


class Styles(object):

    # See: ISO/IEC29500-1:2016 section 18.8.30
    _builtin_formats = {
        0: Format(None),
        1: Format('0'),
        2: Format('0.00'),
        3: Format('#,##0'),
        4: Format('#,##0.00'),
        9: Format('0%'),
        10: Format('0.00%'),
        11: Format('0.00E+00'),
        12: Format('# ?/?'),
        13: Format('# ??/??'),
        14: Format('mm-dd-yy'),
        15: Format('d-mmm-yy'),
        16: Format('d-mmm'),
        17: Format('mmm-yy'),
        18: Format('h:mm AM/PM'),
        19: Format('h:mm:ss AM/PM'),
        20: Format('h:mm'),
        21: Format('h:mm:ss'),
        22: Format('m/d/yy h:mm'),
        37: Format('#,##0;(#,##0)'),
        38: Format('#,##0;[Red](#,##0)'),
        39: Format('#,##0.00;(#,##0.00)'),
        40: Format('#,##0.00;[Red](#,##0.00)'),
        45: Format('mm:ss'),
        46: Format('[h]:mm:ss'),
        47: Format('mmss.0'),
        48: Format('##0.0E+0'),
        49: Format('@')
    }

    def __init__(self, fp):
        self._fp = fp
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self._formats = dict()
        self._cell_style_xfs = list()
        self._cell_xfs = list()

        self._fp.seek(0, os.SEEK_SET)
        reader = RecordReader(self._fp)
        for rectype, rec in reader:
            if rectype == rt.BEGIN_CELL_STYLE_XFS:
                self._cell_style_xfs = [None] * rec.count
                i = 0
                for rectype, rec in reader:
                    if rectype == rt.XF:
                        self._cell_style_xfs[i] = rec
                        i += 1
                    elif rectype == rt.END_CELL_STYLE_XFS:
                        break
            elif rectype == rt.BEGIN_CELL_XFS:
                self._cell_xfs = [None] * rec.count
                i = 0
                for rectype, rec in reader:
                    if rectype == rt.XF:
                        self._cell_xfs[i] = rec
                        i += 1
                    elif rectype == rt.END_CELL_XFS:
                        break
            elif rectype == rt.FMT:
                self._formats[rec.fmtId] = Format(rec.fmtCode)
            elif rectype == rt.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        # TODO
        return None

    def _get_format(self, idx):
        if idx in self._cell_xfs:
            fmt_id = self._cell_xfs[idx].numFmtId
            if fmt_id in self._formats:
                return self._formats[fmt_id]
            elif fmt_id in self._builtin_formats:
                return self._builtin_formats[fmt_id]
        return self._builtin_formats[0]

    def close(self):
        self._fp.close()
