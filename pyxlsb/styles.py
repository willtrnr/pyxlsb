import os
from . import recordtypes as rt
from .recordreader import RecordReader
from .records import FormatRecord


class Format(object):
    __slots__ = ('code', 'is_builtin', '_is_date_format')

    def __init__(self, code):
        self.code = code
        self.is_builtin = False
        self._is_date_format = None

    @property
    def is_date_format(self):
        if self._is_date_format is None:
            self._is_date_format = False
            if self.code is not None:
                # TODO Implement an actual parser
                in_color = 0
                for c in self.code:
                    if c == '[':
                        in_color += 1
                    elif c == ']' and in_color > 0:
                        in_color -= 1
                    elif in_color > 0:
                        continue
                    elif c in ('y', 'm', 'd', 'h', 's'):
                        self._is_date_format = True
                        break
        return self._is_date_format

    def __repr__(self):
        return 'Format(code={}, is_builtin={})' \
            .format(self.code, self.is_builtin)


class BuiltinFormat(Format):
    __slots__ = ()

    def __init__(self, *args, **kwargs):
        super(BuiltinFormat, self).__init__(*args, **kwargs)
        self.is_builtin = True


class Styles(object):

    _general_format = BuiltinFormat(None)

    # See: ISO/IEC29500-1:2016 section 18.8.30
    _builtin_formats = {
        1: BuiltinFormat('0'),
        2: BuiltinFormat('0.00'),
        3: BuiltinFormat('#,##0'),
        4: BuiltinFormat('#,##0.00'),
        9: BuiltinFormat('0%'),
        10: BuiltinFormat('0.00%'),
        11: BuiltinFormat('0.00E+00'),
        12: BuiltinFormat('# ?/?'),
        13: BuiltinFormat('# ??/??'),
        14: BuiltinFormat('mm-dd-yy'),
        15: BuiltinFormat('d-mmm-yy'),
        16: BuiltinFormat('d-mmm'),
        17: BuiltinFormat('mmm-yy'),
        18: BuiltinFormat('h:mm AM/PM'),
        19: BuiltinFormat('h:mm:ss AM/PM'),
        20: BuiltinFormat('h:mm'),
        21: BuiltinFormat('h:mm:ss'),
        22: BuiltinFormat('m/d/yy h:mm'),
        37: BuiltinFormat('#,##0;(#,##0)'),
        38: BuiltinFormat('#,##0;[Red](#,##0)'),
        39: BuiltinFormat('#,##0.00;(#,##0.00)'),
        40: BuiltinFormat('#,##0.00;[Red](#,##0.00)'),
        45: BuiltinFormat('mm:ss'),
        46: BuiltinFormat('[h]:mm:ss'),
        47: BuiltinFormat('mmss.0'),
        48: BuiltinFormat('##0.0E+0'),
        49: BuiltinFormat('@')
    }

    def __init__(self, fp):
        self._fp = fp
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self._formats = {}
        self._cell_style_xfs = []
        self._cell_xfs = []

        self._fp.seek(0, os.SEEK_SET)
        reader = RecordReader(self._fp)
        for rectype, rec in reader:
            if rectype == rt.FMT:
                self._formats[rec.fmtId] = Format(rec.fmtCode)
            elif rectype == rt.BEGIN_CELL_STYLE_XFS:
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
            elif rectype == rt.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        # TODO
        del idx

    def _get_format(self, idx):
        if idx < len(self._cell_xfs):
            fmt_id = self._cell_xfs[idx].numFmtId
            if fmt_id in self._formats:
                return self._formats[fmt_id]
            elif fmt_id in self._builtin_formats:
                return self._builtin_formats[fmt_id]
        return self._general_format

    def close(self):
        self._fp.close()
