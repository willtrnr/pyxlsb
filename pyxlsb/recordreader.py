import sys
import logging
from io import BytesIO
from . import records as recs
from . import recordtypes as rt
from .util import _hexdump
from .datareader import DataReader

if sys.version_info > (3,):
    xrange = range

_logger = logging.getLogger(__name__)


class RecordReader(object):
    _default_record = recs.UnknownRecord

    _records = {
        # Workbook part handlers
        rt.BEGIN_BOOK:       recs.SimpleRecord('Workbook'),
        rt.END_BOOK:         recs.SimpleRecord('WorkbookEnd'),
        rt.WB_PROP:          recs.WorkbookPropertiesRecord,
        rt.BEGIN_BUNDLE_SHS: recs.SimpleRecord('Sheets'),
        rt.END_BUNDLE_SHS:   recs.SimpleRecord('SheetsEnd'),
        rt.BUNDLE_SH:        recs.SheetRecord,

        # Worksheet part handlers
        rt.BEGIN_SHEET:      recs.SimpleRecord('Worksheet'),
        rt.END_SHEET:        recs.SimpleRecord('WorksheetEnd'),
        rt.WS_DIM:           recs.DimensionRecord,
        rt.BEGIN_COL_INFOS:  recs.SimpleRecord('Cols'),
        rt.END_COL_INFOS:    recs.SimpleRecord('ColsEnd'),
        rt.COL_INFO:         recs.ColumnRecord,
        rt.BEGIN_SHEET_DATA: recs.SimpleRecord('SheetData'),
        rt.END_SHEET_DATA:   recs.SimpleRecord('SheetDataEnd'),
        rt.ROW_HDR:          recs.RowRecord,
        rt.CELL_BLANK:       recs.CellRecord,
        rt.CELL_RK:          recs.CellRecord,
        rt.CELL_ERROR:       recs.CellRecord,
        rt.CELL_BOOL:        recs.CellRecord,
        rt.CELL_REAL:        recs.CellRecord,
        rt.CELL_ST:          recs.CellRecord,
        rt.CELL_ISST:        recs.CellRecord,
        rt.FMLA_STRING:      recs.FormulaCellRecord,
        rt.FMLA_NUM:         recs.FormulaCellRecord,
        rt.FMLA_BOOL:        recs.FormulaCellRecord,
        rt.FMLA_ERROR:       recs.FormulaCellRecord,
        rt.H_LINK:           recs.HyperlinkRecord,

        # SharedStrings part handlers
        rt.BEGIN_SST: recs.StringTableRecord,
        rt.END_SST:   recs.SimpleRecord('StringTableEnd'),
        rt.SST_ITEM:  recs.StringTableItemRecord,

        # Styles part handlers
        rt.BEGIN_STYLE_SHEET:    recs.SimpleRecord('StyleSheet'),
        rt.END_STYLE_SHEET:      recs.SimpleRecord('StyleSheetEnd'),
        rt.BEGIN_COLOR_PALETTE:  recs.ColorsRecord,
        rt.END_COLOR_PALETTE:    recs.SimpleRecord('ColorsEnd'),
        rt.BEGIN_DXFS:           recs.DxfsRecord,
        rt.END_DXFS:             recs.SimpleRecord('DxfsEnd'),
        rt.BEGIN_TABLE_STYLES:   recs.TableStylesRecord,
        rt.END_TABLE_STYLES:     recs.SimpleRecord('TableStylesEnd'),
        rt.BEGIN_FILLS:          recs.FillsRecord,
        rt.END_FILLS:            recs.SimpleRecord('FillsEnd'),
        rt.FILL:                 recs.SimpleRecord('Fill'),
        rt.BEGIN_FONTS:          recs.FontsRecord,
        rt.END_FONTS:            recs.SimpleRecord('FontsEnd'),
        rt.FONT:                 recs.FontRecord,
        rt.BEGIN_BORDERS:        recs.BordersRecord,
        rt.END_BORDERS:          recs.SimpleRecord('BordersEnd'),
        rt.BORDER:               recs.SimpleRecord('Border'),
        rt.BEGIN_FMTS:           recs.SimpleRecord('Fmts'),
        rt.END_FMTS:             recs.SimpleRecord('FmtsEnd'),
        rt.FMT:                  recs.FormatRecord,
        rt.BEGIN_CELL_XFS:       recs.CellXfsRecord,
        rt.END_CELL_XFS:         recs.SimpleRecord('CellXfsEnd'),
        rt.XF:                   recs.XfRecord,
        rt.BEGIN_STYLES:         recs.CellStylesRecord,
        rt.END_STYLES:           recs.SimpleRecord('CellStylesEnd'),
        rt.STYLE:                recs.CellStyleRecord,
        rt.BEGIN_CELL_STYLE_XFS: recs.CellStyleXfsRecord,
        rt.END_CELL_STYLE_XFS:   recs.SimpleRecord('CellStyleXfsEnd')
    }

    def __init__(self, fp, enc=None):
        if isinstance(fp, RecordReader):
            self._fp = fp._fp
        elif hasattr(fp, 'read') and hasattr(fp, 'seek'):
            self._fp = fp
        else:
            self._fp = BytesIO(fp)
        self._enc = enc

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def _read_type(self):
        value = self._fp.read(1)
        if not value:
            return None
        value = ord(value)
        if value & 0x80 == 0x80:
            b = self._fp.read(1)
            if not b:
                return None
            value = (value & 0x7F) | ((ord(b) & 0x7F) << 7)
        return value

    def _read_len(self):
        value = 0
        i = 0
        while i < 4:
            b = self._fp.read(1)
            if not b:
                return None
            b = ord(b)
            value |= (b & 0x7F) << (7 * i)
            if b & 0x80 == 0:
                break
            i += 1
        return value

    def next(self):
        rectype = self._read_type()
        if rectype is None:
            raise StopIteration

        reclen = self._read_len()
        if reclen is None:
            raise RuntimeError('incomplete record with type: {}'.format(rectype))

        data = self._fp.read(reclen)

        if _logger.isEnabledFor(logging.DEBUG):
            _logger.debug('Reading record: %s (%d)\n%s', rt._by_num.get(rectype), rectype,
                          _hexdump(data))

        cls = self._records.get(rectype, self._default_record)
        res = cls.read(DataReader(data, enc=self._enc), rectype, reclen)

        _logger.debug('Read result: %s', res)

        return (rectype, res)
