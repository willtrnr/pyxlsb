# noqa: F405
import os
import sys
from . import records as r
from . import recordhandlers as rh
from .datareader import DataReader

if sys.version_info > (3,):
    xrange = range


class RecordReader(DataReader):
    _default_handler = rh.RecordHandler()

    _handlers = {
        # Workbook part handlers
        r.BEGIN_BOOK:       rh.BasicRecordHandler('workbook'),
        r.END_BOOK:         rh.BasicRecordHandler('/workbook'),
        r.WB_PROP:          rh.WorkbookPropertiesHandler(),
        r.BEGIN_BUNDLE_SHS: rh.BasicRecordHandler('sheets'),
        r.END_BUNDLE_SHS:   rh.BasicRecordHandler('/sheets'),
        r.BUNDLE_SH:        rh.SheetHandler(),

        # Worksheet part handlers
        r.BEGIN_SHEET:      rh.BasicRecordHandler('worksheet'),
        r.END_SHEET:        rh.BasicRecordHandler('/worksheet'),
        r.WS_DIM:           rh.DimensionHandler(),
        r.BEGIN_COL_INFOS:  rh.BasicRecordHandler('cols'),
        r.END_COL_INFOS:    rh.BasicRecordHandler('/cols'),
        r.COL_INFO:         rh.ColumnHandler(),
        r.BEGIN_SHEET_DATA: rh.BasicRecordHandler('sheetData'),
        r.END_SHEET_DATA:   rh.BasicRecordHandler('/sheetData'),
        r.ROW_HDR:          rh.RowHandler(),
        r.CELL_BLANK:       rh.CellHandler(),
        r.CELL_RK:          rh.CellHandler(),
        r.CELL_ERROR:       rh.CellHandler(),
        r.CELL_BOOL:        rh.CellHandler(),
        r.CELL_REAL:        rh.CellHandler(),
        r.CELL_ST:          rh.CellHandler(),
        r.CELL_ISST:        rh.CellHandler(),
        r.FMLA_STRING:      rh.FormulaCellHandler(),
        r.FMLA_NUM:         rh.FormulaCellHandler(),
        r.FMLA_BOOL:        rh.FormulaCellHandler(),
        r.FMLA_ERROR:       rh.FormulaCellHandler(),
        r.H_LINK:           rh.HyperlinkHandler(),

        # SharedStrings part handlers
        r.BEGIN_SST: rh.StringTableHandler(),
        r.END_SST:   rh.BasicRecordHandler('/sst'),
        r.SST_ITEM:  rh.StringInstanceHandler(),

        # Styles part handlers
        r.BEGIN_STYLE_SHEET:    rh.BasicRecordHandler('styleSheet'),
        r.END_STYLE_SHEET:      rh.BasicRecordHandler('/styleSheet'),
        r.BEGIN_COLOR_PALETTE:  rh.ColorsHandler(),
        r.END_COLOR_PALETTE:    rh.BasicRecordHandler('/colors'),
        r.BEGIN_DXFS:           rh.DxfsHandler(),
        r.END_DXFS:             rh.BasicRecordHandler('/dxfs'),
        r.BEGIN_TABLE_STYLES:   rh.TableStylesHandler(),
        r.END_TABLE_STYLES:     rh.BasicRecordHandler('/tableStyles'),
        r.BEGIN_FILLS:          rh.FillsHandler(),
        r.END_FILLS:            rh.BasicRecordHandler('/fills'),
        r.FILL:                 rh.BasicRecordHandler('fill'),
        r.BEGIN_FONTS:          rh.FontsHandler(),
        r.END_FONTS:            rh.BasicRecordHandler('/fonts'),
        r.FONT:                 rh.FontHandler(),
        r.BEGIN_BORDERS:        rh.BordersHandler(),
        r.END_BORDERS:          rh.BasicRecordHandler('/borders'),
        r.BORDER:               rh.BasicRecordHandler('border'),
        r.BEGIN_CELL_XFS:       rh.CellXfsHandler(),
        r.END_CELL_XFS:         rh.BasicRecordHandler('/cellXfs'),
        r.XF:                   rh.XfHandler(),
        r.BEGIN_STYLES:         rh.CellStylesHandler(),
        r.END_STYLES:           rh.BasicRecordHandler('/cellStyles'),
        r.STYLE:                rh.CellStyleHandler(),
        r.BEGIN_CELL_STYLE_XFS: rh.CellStyleXfsHandler(),
        r.END_CELL_STYLE_XFS:   rh.BasicRecordHandler('/cellStyleXfs')
    }

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def _read_type(self):
        value = self.read_byte()
        if value is None:
            return None
        if value & 0x80 == 0x80:
            hi = self.read_byte()
            if hi is None:
                return None
            value = (value & 0x7F) | ((hi & 0x7F) << 7)
        return value

    def _read_len(self):
        value = 0
        for i in xrange(4):
            byte = self.read_byte()
            if byte is None:
                return None
            value |= (byte & 0x7F) << (7 * i)
            if byte & 0x80 == 0:
                break
        return value

    def next(self):
        while True:
            rectype = self._read_type()
            if rectype is None:
                raise StopIteration

            reclen = self._read_len()
            if reclen is None:
                raise StopIteration

            handler = self._handlers.get(rectype, self._default_handler)

            boundary = self.tell() + reclen
            res = handler.read(self, rectype, reclen)

            pos = self.tell()
            if pos > boundary:
                raise AssertionError('misbehaving handler ' + str(handler))
            elif pos < boundary:
                self.seek(boundary - pos, os.SEEK_CUR)

            if res is not None:
                return (rectype, res)
