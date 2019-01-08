from enum import Enum
from . import recordtypes as rt
from collections import namedtuple


class BaseRecord(object):
    def __repr__(self):
        args = ('{}={}'.format(str(k), repr(v)) for k, v in self.__dict__.items())
        return '{}({})'.format(self.__class__.__name__, ', '.join(args))

    @classmethod
    def read(cls, reader, rectype, reclen):
        return cls()

    def write(self, writer):
        # TODO Eventually, some day
        pass


def SimpleRecord(name):
    @classmethod
    def read(cls, reader, rectype, reclen):
        res = cls()
        res.brt = rectype
        return res
    return type(name + 'Record', (BaseRecord,), {'brt': 0xFFFF, 'read': read})


class UnknownRecord(BaseRecord):
    brt = 0xFFFF

    def __init__(self, rectype):
        self.brt = rectype

    @classmethod
    def read(cls, reader, rectype, reclen):
        return cls(rectype)


class WorkbookPropertiesRecord(BaseRecord):
    brt = rt.WB_PROP

    def __init__(self, date1904, defaultThemeVersion, codeName):
        self.date1904 = date1904
        self.defaultThemeVersion = defaultThemeVersion
        self.codeName = codeName

    @classmethod
    def read(cls, reader, rectype, reclen):
        flags = reader.read_int()
        theme = reader.read_int()
        name = reader.read_string()
        return cls(flags & 0x01 == 0x01, theme, name)


class SheetState(Enum):
    VISIBLE = 0
    HIDDEN = 1
    VERYHIDDEN = 2


class SheetRecord(BaseRecord):
    brt = rt.BUNDLE_SH

    def __init__(self, state, sheetId, rId, name):
        self.state = state
        self.sheetId = sheetId
        self.rId = rId
        self.name = name

    @classmethod
    def read(cls, reader, rectype, reclen):
        state = reader.read_int()
        sheetid = reader.read_int()
        rid = reader.read_string()
        name = reader.read_string()
        return cls(SheetState(state), sheetid, rid, name)


class DimensionRecord(BaseRecord):
    brt = rt.WS_DIM

    def __init__(self, r, c, h, w):
        self.r = r
        self.c = c
        self.h = h
        self.w = w

    @classmethod
    def read(cls, reader, rectype, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        return cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1)


class ColumnRecord(BaseRecord):
    brt = rt.COL_INFO

    def __init__(self, c1, c2, width, style, customWidth):
        self.c1 = c1
        self.c2 = c2
        self.width = width
        self.style = style
        self.customWidth = customWidth

    @classmethod
    def read(cls, reader, rectype, reclen):
        c1 = reader.read_int()
        c2 = reader.read_int()
        width = reader.read_int() / 256
        style = reader.read_int()
        flags = reader.read_short()
        return cls(c1, c2, width, style, flags & 0x0002 == 0x0002)


class RowRecord(BaseRecord):
    brt = rt.ROW_HDR

    def __init__(self, r):
        self.r = r

    @classmethod
    def read(cls, reader, rectype, reclen):
        r = reader.read_int()
        return cls(r)


# TODO Map error values
class ErrorValue(object):
    __slots__ = ('value',)

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return 'ErrorValue(value=0x{})'.format(hex(self.value))

    def __str__(self):
        return '#ERR!' + str(self.value)


class CellRecord(BaseRecord):
    brt = rt.CELL_BLANK

    def __init__(self, c, v, f, style):
        self.c = c
        self.v = v
        self.f = f
        self.style = style

    @classmethod
    def read(cls, reader, rectype, reclen):
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

        res = cls(col, value, None, style)
        res.brt = rectype
        return res


class FormulaCellRecord(CellRecord):
    brt = rt.FMLA_STRING

    @classmethod
    def read(cls, reader, rectype, reclen):
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
                formula = buf

        res = cls(col, value, formula, style)
        res.brt = rectype
        return res


class HyperlinkRecord(BaseRecord):
    brt = rt.H_LINK

    cls = namedtuple('hyperlink', ['r', 'c', 'h', 'w', 'rId'])

    def __init__(self, r, c, h, w, rId):
        self.r = r
        self.c = c
        self.h = h
        self.w = w
        self.rId = rId

    @classmethod
    def read(cls, reader, rectype, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        rid = reader.read_string()
        return cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1, rid)


class StringTableRecord(BaseRecord):
    brt = rt.BEGIN_SST

    def __init__(self, count, uniqueCount):
        self.count = count
        self.uniqueCount = uniqueCount

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        unique = reader.read_int()
        return cls(count, unique)


class StringTableItemRecord(BaseRecord):
    brt = rt.SST_ITEM

    def __init__(self, t):
        self.t = t

    @classmethod
    def read(cls, reader, rectype, reclen):
        reader.skip(1)
        value = reader.read_string()
        return cls(value)


class ColorsRecord(BaseRecord):
    brt = rt.BEGIN_COLOR_PALETTE

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class DxfsRecord(BaseRecord):
    brt = rt.BEGIN_DXFS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class TableStylesRecord(BaseRecord):
    brt = rt.BEGIN_TABLE_STYLES

    def __init__(self, count, defaultTableStyle, defaultPivotStyle):
        self.count = count
        self.defaultTableStyle = defaultTableStyle
        self.defaultPivotStyle = defaultPivotStyle

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        table = reader.read_string()
        pivot = reader.read_string()
        return cls(count, table, pivot)


class FillsRecord(BaseRecord):
    brt = rt.BEGIN_FILLS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class FontsRecord(BaseRecord):
    brt = rt.BEGIN_FONTS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class BordersRecord(BaseRecord):
    brt = rt.BEGIN_BORDERS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class CellXfsRecord(BaseRecord):
    brt = rt.BEGIN_CELL_XFS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class CellStylesRecord(BaseRecord):
    brt = rt.BEGIN_STYLES

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class CellStyleXfsRecord(BaseRecord):
    brt = rt.BEGIN_CELL_STYLE_XFS

    def __init__(self, count):
        self.count = count

    @classmethod
    def read(cls, reader, rectype, reclen):
        count = reader.read_int()
        return cls(count)


class FontRecord(BaseRecord):
    brt = rt.FONT

    def __init__(self, family):
        self.family = family

    @classmethod
    def read(cls, reader, rectype, reclen):
        reader.skip(21)  # TODO
        family = reader.read_string()
        return cls(family)


class XfRecord(BaseRecord):
    brt = rt.XF

    def __init__(self, numFmtId, fontId, fillId, borderId, xfId):
        self.numFmtId = numFmtId
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.xfId = xfId

    @classmethod
    def read(cls, reader, rectype, reclen):
        # TODO: Speculative and seems wrong
        xf = reader.read_short()
        if xf == 0xFFFF:
            xf = None
        fmtid = reader.read_short()
        font = reader.read_short()
        fill = reader.read_short()
        border = reader.read_short()
        return cls(fmtid, font, fill, border, xf)


class CellStyleRecord(BaseRecord):
    brt = rt.STYLE

    def __init__(self, name):
        self.name = name

    @classmethod
    def read(cls, reader, rectype, reclen):
        reader.skip(8)  # TODO
        name = reader.read_string()
        return cls(name)
