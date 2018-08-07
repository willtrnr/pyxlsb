# noqa: F405
import os
import struct
from . import records
from . import recordhandlers as rh
from .datareader import DataReader

_uint8_t = struct.Struct('<B')


class RecordReader(object):
    default_handler = rh.RecordHandler()

    handlers = {
        # Workbook part handlers
        records.WORKBOOK:     rh.BasicRecordHandler('workbook'),
        records.WORKBOOK_END: rh.BasicRecordHandler('/workbook'),
        records.WORKBOOKPR:   rh.WorkbookPropertiesHandler(),
        records.SHEETS:       rh.BasicRecordHandler('sheets'),
        records.SHEETS_END:   rh.BasicRecordHandler('/sheets'),
        records.SHEET:        rh.SheetHandler(),

        # Worksheet part handlers
        records.WORKSHEET:       rh.BasicRecordHandler('worksheet'),
        records.WORKSHEET_END:   rh.BasicRecordHandler('/worksheet'),
        records.DIMENSION:       rh.DimensionHandler(),
        records.SHEETDATA:       rh.BasicRecordHandler('sheetData'),
        records.SHEETDATA_END:   rh.BasicRecordHandler('/sheetData'),
        records.COLS:            rh.BasicRecordHandler('cols'),
        records.COLS_END:        rh.BasicRecordHandler('/cols'),
        records.COL:             rh.ColumnHandler(),
        records.ROW:             rh.RowHandler(),
        records.BLANK:           rh.CellHandler(),
        records.NUM:             rh.CellHandler(),
        records.BOOLERR:         rh.CellHandler(),
        records.BOOL:            rh.CellHandler(),
        records.FLOAT:           rh.CellHandler(),
        records.STRING:          rh.CellHandler(),
        records.FORMULA_STRING:  rh.FormulaCellHandler(),
        records.FORMULA_FLOAT:   rh.FormulaCellHandler(),
        records.FORMULA_BOOL:    rh.FormulaCellHandler(),
        records.FORMULA_BOOLERR: rh.FormulaCellHandler(),
        records.HYPERLINK:       rh.HyperlinkHandler(),

        # SharedStrings part handlers
        records.SST:     rh.StringTableHandler(),
        records.SST_END: rh.BasicRecordHandler('/sst'),
        records.SI:      rh.StringInstanceHandler(),

        # Styles part handlers
        records.STYLESHEET:       rh.BasicRecordHandler('styleSheet'),
        records.STYLESHEET_END:   rh.BasicRecordHandler('/styleSheet'),
        records.COLORS:           rh.ColorsHandler(),
        records.COLORS_END:       rh.BasicRecordHandler('/colors'),
        records.DXFS:             rh.DxfsHandler(),
        records.DXFS_END:         rh.BasicRecordHandler('/dxfs'),
        records.TABLESTYLES:      rh.TableStylesHandler(),
        records.TABLESTYLES_END:  rh.BasicRecordHandler('/tableStyles'),
        records.FILLS:            rh.FillsHandler(),
        records.FILLS_END:        rh.BasicRecordHandler('/fills'),
        records.FONTS:            rh.FontsHandler(),
        records.FONTS_END:        rh.BasicRecordHandler('/fonts'),
        records.BORDERS:          rh.BordersHandler(),
        records.BORDERS_END:      rh.BasicRecordHandler('/borders'),
        records.CELLXFS:          rh.CellXfsHandler(),
        records.CELLXFS_END:      rh.BasicRecordHandler('/cellXfs'),
        records.CELLSTYLES:       rh.CellStylesHandler(),
        records.CELLSTYLES_END:   rh.BasicRecordHandler('/cellStyles'),
        records.CELLSTYLEXFS:     rh.CellStyleXfsHandler(),
        records.CELLSTYLEXFS_END: rh.BasicRecordHandler('/cellStyleXfs'),
        records.FONT:             rh.FontHandler(),
        records.XF:               rh.XfHandler(),
        records.CELLSTYLE:        rh.CellStyleHandler()
    }

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

    def register_handler(self, recid, handler):
        self.handlers[recid] = handler

    def next(self):
        ret = None
        while ret is None:
            if self._debug:
                pos = self._fp.tell()
            recid = self.read_id()
            reclen = self.read_len()
            if recid is None or reclen is None:
                raise StopIteration
            recdata = self._fp.read(reclen)
            ret = self.handlers.get(recid, self.default_handler).read(DataReader(recdata), recid, reclen)
            if self._debug:
                print('{:08X}  {:04X}  {:<6} {}\n{}'.format(pos, recid, reclen, ' '.join('{:02X}'.format(b) for b in recdata), ret))
        return (recid, ret)

    def close(self):
        self._fp.close()
