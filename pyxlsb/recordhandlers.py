from . import records
from .formula import Formula
from collections import namedtuple

class RecordHandler(object):
    def read(self, reader, recid, reclen):
        if reclen > 0:
            reader.skip(reclen)

    def write(self, writer, data):
        # TODO Eventually some day
        pass


class BasicRecordHandler(RecordHandler):
    def __init__(self, name=None):
        super(BasicRecordHandler, self).__init__()
        self.name = name

    def read(self, reader, recid, reclen):
        super(BasicRecordHandler, self).read(reader, recid, reclen)
        reader.skip(reclen)
        return self.name


class WorkbookPropertiesHandler(RecordHandler):
    cls = namedtuple('workbookPr', ['date1904', 'defaultThemeVersion'])

    def read(self, reader, recid, reclen):
        flags = reader.read_short() # TODO: This contains the 1904 flag
        reader.skip(2) # Not sure what this is, other flags probably
        theme = reader.read_int()
        reader.skip(4) # Also not sure, more flags?
        return self.cls(flags & 0x01 == 0x01, theme)


class SheetHandler(RecordHandler):
    cls = namedtuple('sheet', ['sheetId', 'rId', 'name'])

    def read(self, reader, recid, reclen):
        reader.skip(4)
        sheetid = reader.read_int()
        rid = reader.read_string()
        name = reader.read_string()
        return self.cls(sheetid, rid, name)


class DimensionHandler(RecordHandler):
    cls = namedtuple('dimension', ['r', 'c', 'h', 'w'])

    def read(self, reader, recid, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1)


class ColumnHandler(RecordHandler):
    cls = namedtuple('col', ['c1', 'c2', 'width', 'style', 'customWidth'])

    def read(self, reader, recid, reclen):
        c1 = reader.read_int()
        c2 = reader.read_int()
        width = reader.read_int() / 256
        style = reader.read_int()
        flags = reader.read_short()
        return self.cls(c1, c2, width, style, flags & 0x0002 == 0x0002)


class RowHandler(RecordHandler):
    cls = namedtuple('row', ['r'])

    def read(self, reader, recid, reclen):
        r = reader.read_int()
        return self.cls(r)


class CellHandler(RecordHandler):
    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def read(self, reader, recid, reclen):
        col = reader.read_int()
        style = reader.read_int()

        value = None
        if recid == records.NUM:
            value = reader.read_rk()
        elif recid == records.BOOLERR:
            # TODO Map error values
            value = hex(reader.read_byte())
        elif recid == records.BOOL:
            value = reader.read_bool()
        elif recid == records.FLOAT:
            value = reader.read_double()
        elif recid == records.STRING:
            value = reader.read_int()

        return self.cls(col, value, None, style)

class FormulaCellHandler(RecordHandler):
    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def read(self, reader, recid, reclen):
        col = reader.read_int()
        style = reader.read_int()

        value = None
        if recid == records.FORMULA_STRING:
            value = reader.read_string()
        elif recid == records.FORMULA_FLOAT:
            value = reader.read_double()
        elif recid == records.FORMULA_BOOL:
            value = reader.read_bool()
        elif recid == records.FORMULA_BOOLERR:
            # TODO Map error values
            value = hex(reader.read_byte())

        formula = None
        # 0x0001 = Recalc always, 0x0002 = Calc on open, 0x0008 = Part of shared
        grbits = reader.read_short()
        sz = reader.read_int()
        if sz:
            buf = reader.read(sz)
            if len(buf) == sz:
                formula = Formula.parse(buf)

        return self.cls(col, value, formula, style)


class HyperlinkHandler(RecordHandler):
    cls = namedtuple('hyperlink', ['r', 'c', 'h', 'w', 'rId'])

    def read(self, reader, recid, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        rid = reader.read_string()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1, rid)


class StringTableHandler(RecordHandler):
    cls = namedtuple('sst', ['count', 'uniqueCount'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        unique = reader.read_int()
        return self.cls(count, unique)


class StringInstanceHandler(RecordHandler):
    cls = namedtuple('si', ['t'])

    def read(self, reader, recid, reclen):
        reader.skip(1)
        value = reader.read_string()
        return self.cls(value)


class ColorsHandler(RecordHandler):
    cls = namedtuple('colors', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class DxfsHandler(RecordHandler):
    cls = namedtuple('dxfs', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class TableStylesHandler(RecordHandler):
    cls = namedtuple('tableStyles', ['count', 'defaultTableStyle', 'defaultPivotStyle'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        table = reader.read_string()
        pivot = reader.read_string()
        return self.cls(count, table, pivot)


class FillsHandler(RecordHandler):
    cls = namedtuple('fills', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class FontsHandler(RecordHandler):
    cls = namedtuple('fonts', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class BordersHandler(RecordHandler):
    cls = namedtuple('borders', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellXfsHandler(RecordHandler):
    cls = namedtuple('cellXfs', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellStylesHandler(RecordHandler):
    cls = namedtuple('cellStyles', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)


class CellStyleXfsHandler(RecordHandler):
    cls = namedtuple('cellStyleXfs', ['count'])

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        return self.cls(count)

class FontHandler(RecordHandler):
    cls = namedtuple('font', ['family'])

    def read(self, reader, recid, reclen):
        reader.skip(21) # TODO
        family = reader.read_string()
        return self.cls(family)

class XfHandler(RecordHandler):
    cls = namedtuple('xf', ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'])
    def read(self, reader, recid, reclen):
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
    cls = namedtuple('cellStyle', ['name'])

    def read(self, reader, recid, reclen):
        reader.skip(8) # TODO
        name = reader.read_string()
        return self.cls(name)
