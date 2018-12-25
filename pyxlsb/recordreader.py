# noqa: F405
import os
import sys
from . import recordhandlers as rh
from . import recordtypes as rt
from .datareader import DataReader

if sys.version_info > (3,):
    xrange = range


class RecordReader(DataReader):
    _default_handler = rh.RecordHandler()

    _handlers = {
        # Workbook part handlers
        rt.BEGIN_BOOK:       rh.SimpleRecordHandler('workbook'),
        rt.END_BOOK:         rh.SimpleRecordHandler('/workbook'),
        rt.WB_PROP:          rh.WorkbookPropertiesHandler(),
        rt.BEGIN_BUNDLE_SHS: rh.SimpleRecordHandler('sheets'),
        rt.END_BUNDLE_SHS:   rh.SimpleRecordHandler('/sheets'),
        rt.BUNDLE_SH:        rh.SheetHandler(),

        # Worksheet part handlers
        rt.BEGIN_SHEET:      rh.SimpleRecordHandler('worksheet'),
        rt.END_SHEET:        rh.SimpleRecordHandler('/worksheet'),
        rt.WS_DIM:           rh.DimensionHandler(),
        rt.BEGIN_COL_INFOS:  rh.SimpleRecordHandler('cols'),
        rt.END_COL_INFOS:    rh.SimpleRecordHandler('/cols'),
        rt.COL_INFO:         rh.ColumnHandler(),
        rt.BEGIN_SHEET_DATA: rh.SimpleRecordHandler('sheetData'),
        rt.END_SHEET_DATA:   rh.SimpleRecordHandler('/sheetData'),
        rt.ROW_HDR:          rh.RowHandler(),
        rt.CELL_BLANK:       rh.CellHandler(),
        rt.CELL_RK:          rh.CellHandler(),
        rt.CELL_ERROR:       rh.CellHandler(),
        rt.CELL_BOOL:        rh.CellHandler(),
        rt.CELL_REAL:        rh.CellHandler(),
        rt.CELL_ST:          rh.CellHandler(),
        rt.CELL_ISST:        rh.CellHandler(),
        rt.FMLA_STRING:      rh.FormulaCellHandler(),
        rt.FMLA_NUM:         rh.FormulaCellHandler(),
        rt.FMLA_BOOL:        rh.FormulaCellHandler(),
        rt.FMLA_ERROR:       rh.FormulaCellHandler(),
        rt.H_LINK:           rh.HyperlinkHandler(),

        # SharedStrings part handlers
        rt.BEGIN_SST: rh.StringTableHandler(),
        rt.END_SST:   rh.SimpleRecordHandler('/sst'),
        rt.SST_ITEM:  rh.StringTableItemHandler(),

        # Styles part handlers
        rt.BEGIN_STYLE_SHEET:    rh.SimpleRecordHandler('styleSheet'),
        rt.END_STYLE_SHEET:      rh.SimpleRecordHandler('/styleSheet'),
        rt.BEGIN_COLOR_PALETTE:  rh.ColorsHandler(),
        rt.END_COLOR_PALETTE:    rh.SimpleRecordHandler('/colors'),
        rt.BEGIN_DXFS:           rh.DxfsHandler(),
        rt.END_DXFS:             rh.SimpleRecordHandler('/dxfs'),
        rt.BEGIN_TABLE_STYLES:   rh.TableStylesHandler(),
        rt.END_TABLE_STYLES:     rh.SimpleRecordHandler('/tableStyles'),
        rt.BEGIN_FILLS:          rh.FillsHandler(),
        rt.END_FILLS:            rh.SimpleRecordHandler('/fills'),
        rt.FILL:                 rh.SimpleRecordHandler('fill'),
        rt.BEGIN_FONTS:          rh.FontsHandler(),
        rt.END_FONTS:            rh.SimpleRecordHandler('/fonts'),
        rt.FONT:                 rh.FontHandler(),
        rt.BEGIN_BORDERS:        rh.BordersHandler(),
        rt.END_BORDERS:          rh.SimpleRecordHandler('/borders'),
        rt.BORDER:               rh.SimpleRecordHandler('border'),
        rt.BEGIN_CELL_XFS:       rh.CellXfsHandler(),
        rt.END_CELL_XFS:         rh.SimpleRecordHandler('/cellXfs'),
        rt.XF:                   rh.XfHandler(),
        rt.BEGIN_STYLES:         rh.CellStylesHandler(),
        rt.END_STYLES:           rh.SimpleRecordHandler('/cellStyles'),
        rt.STYLE:                rh.CellStyleHandler(),
        rt.BEGIN_CELL_STYLE_XFS: rh.CellStyleXfsHandler(),
        rt.END_CELL_STYLE_XFS:   rh.SimpleRecordHandler('/cellStyleXfs')
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
