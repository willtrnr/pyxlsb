# noqa: F405
import os
import struct
from . import records as r
from . import recordhandlers as rh
from .datareader import DataReader

_uint8_t = struct.Struct('<B')


class RecordReader(object):
    default_handler = rh.RecordHandler()

    handlers = [default_handler] * 0x07FF

    # Workbook part handlers
    handlers[r.WORKBOOK]     = rh.BasicRecordHandler('workbook')
    handlers[r.WORKBOOK_END] = rh.BasicRecordHandler('/workbook')
    handlers[r.WORKBOOKPR]   = rh.WorkbookPropertiesHandler()
    handlers[r.SHEETS]       = rh.BasicRecordHandler('sheets')
    handlers[r.SHEETS_END]   = rh.BasicRecordHandler('/sheets')
    handlers[r.SHEET]        = rh.SheetHandler()

    # Worksheet part handlers
    handlers[r.WORKSHEET]       = rh.BasicRecordHandler('worksheet')
    handlers[r.WORKSHEET_END]   = rh.BasicRecordHandler('/worksheet')
    handlers[r.DIMENSION]       = rh.DimensionHandler()
    handlers[r.SHEETDATA]       = rh.BasicRecordHandler('sheetData')
    handlers[r.SHEETDATA_END]   = rh.BasicRecordHandler('/sheetData')
    handlers[r.COLS]            = rh.BasicRecordHandler('cols')
    handlers[r.COLS_END]        = rh.BasicRecordHandler('/cols')
    handlers[r.COL]             = rh.ColumnHandler()
    handlers[r.ROW]             = rh.RowHandler()
    handlers[r.BLANK]           = rh.CellHandler()
    handlers[r.NUM]             = rh.CellHandler()
    handlers[r.BOOLERR]         = rh.CellHandler()
    handlers[r.BOOL]            = rh.CellHandler()
    handlers[r.FLOAT]           = rh.CellHandler()
    handlers[r.STRING]          = rh.CellHandler()
    handlers[r.FORMULA_STRING]  = rh.FormulaCellHandler()
    handlers[r.FORMULA_FLOAT]   = rh.FormulaCellHandler()
    handlers[r.FORMULA_BOOL]    = rh.FormulaCellHandler()
    handlers[r.FORMULA_BOOLERR] = rh.FormulaCellHandler()
    handlers[r.HYPERLINK]       = rh.HyperlinkHandler()

    # SharedStrings part handlers
    handlers[r.SST]     = rh.StringTableHandler()
    handlers[r.SST_END] = rh.BasicRecordHandler('/sst')
    handlers[r.SI]      = rh.StringInstanceHandler()

    # Styles part handlers
    handlers[r.STYLESHEET]       = rh.BasicRecordHandler('styleSheet')
    handlers[r.STYLESHEET_END]   = rh.BasicRecordHandler('/styleSheet')
    handlers[r.COLORS]           = rh.ColorsHandler()
    handlers[r.COLORS_END]       = rh.BasicRecordHandler('/colors')
    handlers[r.DXFS]             = rh.DxfsHandler()
    handlers[r.DXFS_END]         = rh.BasicRecordHandler('/dxfs')
    handlers[r.TABLESTYLES]      = rh.TableStylesHandler()
    handlers[r.TABLESTYLES_END]  = rh.BasicRecordHandler('/tableStyles')
    handlers[r.FILLS]            = rh.FillsHandler()
    handlers[r.FILLS_END]        = rh.BasicRecordHandler('/fills')
    handlers[r.FONTS]            = rh.FontsHandler()
    handlers[r.FONTS_END]        = rh.BasicRecordHandler('/fonts')
    handlers[r.BORDERS]          = rh.BordersHandler()
    handlers[r.BORDERS_END]      = rh.BasicRecordHandler('/borders')
    handlers[r.CELLXFS]          = rh.CellXfsHandler()
    handlers[r.CELLXFS_END]      = rh.BasicRecordHandler('/cellXfs')
    handlers[r.CELLSTYLES]       = rh.CellStylesHandler()
    handlers[r.CELLSTYLES_END]   = rh.BasicRecordHandler('/cellStyles')
    handlers[r.CELLSTYLEXFS]     = rh.CellStyleXfsHandler()
    handlers[r.CELLSTYLEXFS_END] = rh.BasicRecordHandler('/cellStyleXfs')
    handlers[r.FONT]             = rh.FontHandler()
    handlers[r.XF]               = rh.XfHandler()
    handlers[r.CELLSTYLE]        = rh.CellStyleHandler()

    def __init__(self, fp, _debug=False):
        super(RecordReader, self).__init__()
        self._fp = fp
        self._debug = _debug

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def tell(self):
        return self._fp.tell()

    def seek(self, offset, whence=os.SEEK_SET):
        self._fp.seek(offset, whence)

    def read_id(self):
        value = 0
        for i in range(4):
            byte = self._fp.read(1)
            if not byte:
                return None
            byte = _uint8_t.unpack(byte)[0]
            value += byte << 8 * i
            if byte & 0x80 == 0:
                break
        return value

    def read_len(self):
        value = 0
        for i in range(4):
            byte = self._fp.read(1)
            if not byte:
                return None
            byte = _uint8_t.unpack(byte)[0]
            value += (byte & 0x7F) << (7 * i)
            if byte & 0x80 == 0:
                break
        return value

    def next(self):
        ret = None
        while ret is None:
            if self._debug:
                pos = self._fp.tell()
            recid = self.read_id()
            if recid is None:
                raise StopIteration
            reclen = self.read_len()
            if reclen is None:
                raise StopIteration
            recdata = self._fp.read(reclen)
            if recid < len(self.handlers):
                handler = self.handlers[recid]
            else:
                handler = self.default_handler
            ret = handler.read(DataReader(recdata), recid, reclen)
            if self._debug:
                print('{:08X}  {:04X}  {:<6} {}\n{}'.format(pos, recid, reclen, ' '.join('{:02X}'.format(b) for b in recdata), ret))
        return (recid, ret)

    def close(self):
        self._fp.close()
