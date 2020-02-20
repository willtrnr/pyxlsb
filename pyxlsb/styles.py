import os
from . import recordtypes as rt
from .recordreader import RecordReader
from .records import XfRecord, FormatRecord
from .conv import detect_dtype

class Styles(object):
    def __init__(self, fp):
        self._fp = fp
        self._parse()
        self._default_styles = {0: {'format': 'General', 'dtype': ''},
                                             1: {'format': '0', 'dtype': 'float64'},
                                             2: {'format': '0.00', 'dtype': 'float64'},
                                             3: {'format': '#,##0', 'dtype': 'float64'},
                                             4: {'format': '#,##0.00', 'dtype': 'float64'},
                                             5: {'format': '($#,##0_);($#,##0)', 'dtype': 'float64'},
                                             6: {'format': '($#,##0_);[Red]($#,##0)', 'dtype': 'float64'},
                                             7: {'format': '($#,##0.00_);($#,##0.00)', 'dtype': 'float64'},
                                             8: {'format': '($#,##0.00_);[Red]($#,##0.00)', 'dtype': 'float64'},
                                             9: {'format': '0%', 'dtype': 'float64'},
                                             10: {'format': '0.00%', 'dtype': 'float64'},
                                             11: {'format': '0.00E+00', 'dtype': 'float64'},
                                             12: {'format': '# ?/?', 'dtype': 'float64'},
                                             13: {'format': '# ??/??', 'dtype': 'float64'},
                                             14: {'format': 'm/d/yy', 'dtype': 'datetime'},
                                             15: {'format': 'd-mmm-yy', 'dtype': 'datetime'},
                                             16: {'format': 'd-mmm', 'dtype': 'datetime'},
                                             17: {'format': 'mmm-yy', 'dtype': 'datetime'},
                                             18: {'format': 'h:mm AM/PM', 'dtype': 'datetime'},
                                             19: {'format': 'h:mm:ss AM/PM', 'dtype': 'datetime'},
                                             20: {'format': 'h:mm', 'dtype': 'datetime'},
                                             21: {'format': 'h:mm:ss', 'dtype': 'datetime'},
                                             22: {'format': 'm/d/yy h:mm', 'dtype': 'datetime'},
                                             37: {'format': '(#,##0_);(#,##0)', 'dtype': 'float64'},
                                             38: {'format': '(#,##0_);[Red](#,##0)', 'dtype': 'float64'},
                                             39: {'format': '(#,##0.00_);(#,##0.00)', 'dtype': 'float64'},
                                             40: {'format': '(#,##0.00_);[Red](#,##0.00)', 'dtype': 'float64'},
                                             41: {'format': '_(* #,##0_);_(* (#,##0);_(* "-"_);_(_)', 'dtype': 'float64'},
                                             42: {'format': '_($* #,##0_);_($* (#,##0);_($* "-"_);_(_)', 'dtype': 'float64'},
                                             43: {'format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(_)', 'dtype': 'float64'},
                                             44: {'format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(_)', 'dtype': 'float64'},
                                             45: {'format': 'mm:ss', 'dtype': 'datetime'},
                                             46: {'format': '[h]:mm:ss', 'dtype': 'datetime'},
                                             47: {'format': 'mm:ss.0', 'dtype': 'datetime'},
                                             48: {'format': '##0.0E+0', 'dtype': 'float64'},
                                             49: {'format': '@', 'dtype': 'string'}}

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

        self._XfRecord = dict()
        self._FormatRecord = dict()

        self._fp.seek(0, os.SEEK_SET)
        for rectype, rec in RecordReader(self._fp):
            # TODO
            if isinstance(rec, XfRecord):
                self._XfRecord[len(self._XfRecord) - 1] = rec
            elif isinstance(rec, FormatRecord):
                self._FormatRecord[rec.fmtId] = rec
                self._FormatRecord[rec.fmtId].dtype = detect_dtype(self._FormatRecord[rec.fmtId].fmtCode)

            if rectype == rt.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        if idx is None:
            return FormatRecord(fmtId=-1, fmtCode='General', dtype="")
        elif idx in self._XfRecord:
            numFmtId = self._XfRecord[idx].numFmtId
            if numFmtId in self._FormatRecord:
                return self._FormatRecord[numFmtId]
            elif numFmtId in self._default_styles:
                return FormatRecord(fmtId=numFmtId, fmtCode=self._default_styles[numFmtId]["format"], dtype=self._default_styles[numFmtId]["dtype"])
        return FormatRecord(fmtId=-1, fmtCode='General', dtype="")

    def close(self):
        self._fp.close()
