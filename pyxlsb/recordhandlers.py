from enum import Enum
from . import recordtypes as rt
from .formula import Formula
from collections import namedtuple


class RecordHandler(object):
    def read(self, reader, rectype, reclen):
        reader.skip(reclen)

    def write(self, writer, data):
        # TODO Eventually, some day
        writer.write(data)


class SimpleRecordHandler(RecordHandler):
    __slots__ = ['name']

    def __init__(self, name=None):
        self.name = name

    def read(self, reader, rectype, reclen):
        reader.skip(reclen)
        return self.name


class WorkbookPropertiesHandler(RecordHandler):
    brt = rt.WB_PROP

    cls = namedtuple('workbookPr', ['date1904', 'defaultThemeVersion', 'codeName'])

    def read(self, reader, rectype, reclen):
        flags = reader.read_int()
        theme = reader.read_int()
        name = reader.read_string()
        return self.cls(flags & 0x01 == 0x01, theme, name)


class SheetState(Enum):
    VISIBLE = 0
    HIDDEN = 1
    VERYHIDDEN = 2


class SheetHandler(RecordHandler):
    brt = rt.BUNDLE_SH

    cls = namedtuple('sheet', ['state', 'sheetId', 'rId', 'name'])

    def read(self, reader, rectype, reclen):
        state = reader.read_int()
        sheetid = reader.read_int()
        rid = reader.read_string()
        name = reader.read_string()
        return self.cls(SheetState(state), sheetid, rid, name)


class DimensionHandler(RecordHandler):
    brt = rt.WS_DIM

    cls = namedtuple('dimension', ['r', 'c', 'h', 'w'])

    def read(self, reader, rectype, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1)


class ColumnHandler(RecordHandler):
    brt = rt.COL_INFO

    cls = namedtuple('col', ['c1', 'c2', 'width', 'style', 'customWidth'])

    def read(self, reader, rectype, reclen):
        c1 = reader.read_int()
        c2 = reader.read_int()
        width = reader.read_int() / 256
        style = reader.read_int()
        flags = reader.read_short()
        return self.cls(c1, c2, width, style, flags & 0x0002 == 0x0002)


class RowHandler(RecordHandler):
    brt = rt.ROW_HDR

    cls = namedtuple('row', ['r'])

    def read(self, reader, rectype, reclen):
        r = reader.read_int()
        return self.cls(r)


# TODO Map error values
class ErrorValue(object):
    __slots__ = ['value']

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return 'ErrorValue(value=0x{})'.format(hex(self.value))

    def __str__(self):
        return '#ERR!' + str(self.value)


class CellHandler(RecordHandler):
    brt = rt.CELL_BLANK

    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def read(self, reader, rectype, reclen):
        col = reader.read_int()
        style = reader.read_int()

        if rectype == rt.CELL_BLANK:
            value = None
        elif rectype == rt.CELL_RK:
            value = reader.read_rk()
        elif rectype == rt.CELL_ERROR:
            value = ErrorValue(reader.read_byte())
        elif rectype == rt.CELL_BOOL:
            value = reader.read_bool()
        elif rectype == rt.CELL_REAL:
            value = reader.read_double()
        elif rectype == rt.CELL_ST:
            value = reader.read_string()
        elif rectype == rt.CELL_ISST:
            value = reader.read_int()
        else:
            raise ValueError('not a cell record: ' + str(rectype))

        return self.cls(col, value, None, style)


class FormulaCellHandler(RecordHandler):
    brt = rt.FMLA_STRING

    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def read(self, reader, rectype, reclen):
        col = reader.read_int()
        style = reader.read_int()

        if rectype == rt.FMLA_STRING:
            value = reader.read_string()
        elif rectype == rt.FMLA_NUM:
            value = reader.read_double()
        elif rectype == rt.FMLA_BOOL:
            value = reader.read_bool()
        elif rectype == rt.FMLA_ERROR:
            value = ErrorValue(reader.read_byte())
        else:
            raise ValueError('not a formula cell record: ' + str(rectype))

        formula = None
        # 0x0001 = Recalc always, 0x0002 = Calc on open, 0x0008 = Part of shared
        reader.read_short()  # TODO Handle the flags
        sz = reader.read_int()
        if sz:
            buf = reader.read(sz)
            if len(buf) == sz:
                formula = Formula.parse(buf)

        return self.cls(col, value, formula, style)


class HyperlinkHandler(RecordHandler):
    brt = rt.H_LINK

    cls = namedtuple('hyperlink', ['r', 'c', 'h', 'w', 'rId'])

    def read(self, reader, rectype, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        rid = reader.read_string()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1, rid)


class StringTableHandler(RecordHandler):
    brt = rt.BEGIN_SST

    cls = namedtuple('sst', ['count', 'uniqueCount'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        unique = reader.read_int()
        return self.cls(count, unique)


class StringTableItemHandler(RecordHandler):
    brt = rt.SST_ITEM

    cls = namedtuple('si', ['t'])

    def read(self, reader, rectype, reclen):
        reader.skip(1)
        value = reader.read_string()
        return self.cls(value)


class ColorsHandler(RecordHandler):
    brt = rt.BEGIN_COLOR_PALETTE

    cls = namedtuple('colors', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class DxfsHandler(RecordHandler):
    brt = rt.BEGIN_DXFS

    cls = namedtuple('dxfs', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class TableStylesHandler(RecordHandler):
    brt = rt.BEGIN_TABLE_STYLES

    cls = namedtuple('tableStyles', ['count', 'defaultTableStyle', 'defaultPivotStyle'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        table = reader.read_string()
        pivot = reader.read_string()
        return self.cls(count, table, pivot)


class FillsHandler(RecordHandler):
    brt = rt.BEGIN_FILLS

    cls = namedtuple('fills', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class FontsHandler(RecordHandler):
    brt = rt.BEGIN_FONTS

    cls = namedtuple('fonts', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class BordersHandler(RecordHandler):
    brt = rt.BEGIN_BORDERS

    cls = namedtuple('borders', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellXfsHandler(RecordHandler):
    brt = rt.BEGIN_CELL_XFS

    cls = namedtuple('cellXfs', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellStylesHandler(RecordHandler):
    brt = rt.BEGIN_STYLES

    cls = namedtuple('cellStyles', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellStyleXfsHandler(RecordHandler):
    brt = rt.BEGIN_CELL_STYLE_XFS

    cls = namedtuple('cellStyleXfs', ['count'])

    def read(self, reader, rectype, reclen):
        count = reader.read_int()
        return self.cls(count)


class FontHandler(RecordHandler):
    brt = rt.FONT

    cls = namedtuple('font', ['family'])

    def read(self, reader, rectype, reclen):
        reader.skip(21)  # TODO
        family = reader.read_string()
        return self.cls(family)


class XfHandler(RecordHandler):
    brt = rt.XF

    cls = namedtuple('xf', ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'])

    def read(self, reader, rectype, reclen):
        # TODO: Speculative and seems wrong
        xf = reader.read_short()
        if xf == 0xFFFF:
            xf = None
        fmtid = reader.read_short()
        font = reader.read_short()
        fill = reader.read_short()
        border = reader.read_short()
        return self.cls(fmtid, font, fill, border, xf)


class CellStyleHandler(RecordHandler):
    brt = rt.STYLE

    cls = namedtuple('cellStyle', ['name'])

    def read(self, reader, rectype, reclen):
        reader.skip(8)  # TODO
        name = reader.read_string()
        return self.cls(name)
