import os
from . import recordtypes as rt
from .recordreader import RecordReader
from .records import FormatRecord


class Styles(object):

    # from https://github.com/jmcnamara/XlsxWriter/blob/master/xlsxwriter/styles.py
    _builtin_formats = {
        0: FormatRecord(fmtId=0, fmtCode='General'),
        1: FormatRecord(fmtId=1, fmtCode='0'),
        2: FormatRecord(fmtId=2, fmtCode='0.00'),
        3: FormatRecord(fmtId=3, fmtCode='#,##0'),
        4: FormatRecord(fmtId=4, fmtCode='#,##0.00'),
        5: FormatRecord(fmtId=5, fmtCode='($#,##0_);($#,##0)'),
        6: FormatRecord(fmtId=6, fmtCode='($#,##0_);[Red]($#,##0)'),
        7: FormatRecord(fmtId=7, fmtCode='($#,##0.00_);($#,##0.00)'),
        8: FormatRecord(fmtId=8, fmtCode='($#,##0.00_);[Red]($#,##0.00)'),
        9: FormatRecord(fmtId=9, fmtCode='0%'),
        10: FormatRecord(fmtId=10, fmtCode='0.00%'),
        11: FormatRecord(fmtId=11, fmtCode='0.00E+00'),
        12: FormatRecord(fmtId=12, fmtCode='# ?/?'),
        13: FormatRecord(fmtId=13, fmtCode='# ??/??'),
        14: FormatRecord(fmtId=14, fmtCode='m/d/yy'),
        15: FormatRecord(fmtId=15, fmtCode='d-mmm-yy'),
        16: FormatRecord(fmtId=16, fmtCode='d-mmm'),
        17: FormatRecord(fmtId=17, fmtCode='mmm-yy'),
        18: FormatRecord(fmtId=18, fmtCode='h:mm AM/PM'),
        19: FormatRecord(fmtId=19, fmtCode='h:mm:ss AM/PM'),
        20: FormatRecord(fmtId=20, fmtCode='h:mm'),
        21: FormatRecord(fmtId=21, fmtCode='h:mm:ss'),
        22: FormatRecord(fmtId=22, fmtCode='m/d/yy h:mm'),
        37: FormatRecord(fmtId=37, fmtCode='(#,##0_);(#,##0)'),
        38: FormatRecord(fmtId=38, fmtCode='(#,##0_);[Red](#,##0)'),
        39: FormatRecord(fmtId=39, fmtCode='(#,##0.00_);(#,##0.00)'),
        40: FormatRecord(fmtId=40, fmtCode='(#,##0.00_);[Red](#,##0.00)'),
        41: FormatRecord(fmtId=41, fmtCode='_(* #,##0_);_(* (#,##0);_(* "-"_);_(_)'),
        42: FormatRecord(fmtId=42, fmtCode='_($* #,##0_);_($* (#,##0);_($* "-"_);_(_)'),
        43: FormatRecord(fmtId=43, fmtCode='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(_)'),
        44: FormatRecord(fmtId=44, fmtCode='_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(_)'),
        45: FormatRecord(fmtId=45, fmtCode='mm:ss'),
        46: FormatRecord(fmtId=46, fmtCode='[h]:mm:ss'),
        47: FormatRecord(fmtId=47, fmtCode='mm:ss.0'),
        48: FormatRecord(fmtId=48, fmtCode='##0.0E+0'),
        49: FormatRecord(fmtId=49, fmtCode='@')
    }

    def __init__(self, fp):
        self._fp = fp
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self._colors = list()
        self._dxfs = list()
        self._table_styles = list()
        self._fills = list()
        self._fonts = list()
        self._borders = list()
        self._cell_xfs = list()
        self._cell_styles = list()
        self._cell_style_xfs = list()

        self._xf_record = dict()
        self._format_record = dict()

        self._fp.seek(0, os.SEEK_SET)
        for rectype, rec in RecordReader(self._fp):
            # TODO
            if rectype == rt.XF:
                self._xf_record[len(self._xf_record) - 1] = rec
            elif rectype == rt.FMT:
                self._format_record[rec.fmtId] = rec
            elif rectype == rt.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        return None

    def get_format(self, idx):
        if idx in self._xf_record:
            numFmtId = self._xf_record[idx].numFmtId
            if numFmtId in self._format_record:
                return self._format_record[numFmtId]
            elif numFmtId in self._builtin_formats:
                return self._builtin_formats[numFmtId]
        return self._builtin_formats[0]

    def get_dtype(self, idx):
        if idx in self._xf_record:
            numFmtId = self._xf_record[idx].numFmtId
            if numFmtId in self._format_record:
                fmtCode = self._format_record[numFmtId].fmtCode
            elif numFmtId in self._builtin_formats:
                fmtCode = self._builtin_formats[numFmtId].fmtCode
            else:
                return ""
            for char in ["d", "m", "y", "h", "s"]:
                if char in fmtCode:
                    return "datetime"
            for char in ["0.00", "0,00", "#.##", "#,##"]:
                if char in fmtCode:
                    return "float64"
        return ""

    def close(self):
        self._fp.close()
