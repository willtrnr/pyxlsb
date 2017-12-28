from . import records
from .formula import Formula
from collections import namedtuple

class RecordHandler(object):
    def __init__(self):
        super(RecordHandler, self).__init__()

    def read(self, reader, recid, reclen):
        if reclen > 0:
            reader.skip(reclen)

    def write(self, writer, data):
        pass


class BasicRecordHandler(RecordHandler):
    def __init__(self, name=None):
        super(BasicRecordHandler, self).__init__()
        self.name = name

    def read(self, reader, recid, reclen):
        super(BasicRecordHandler, self).read(reader, recid, reclen)
        return self.name


class FileVersionHandler(RecordHandler):
    cls = namedtuple('fileVersion', ['lastEdited', 'lowestEdited', 'rupBuild'])

    def __init__(self):
        super(FileVersionHandler, self).__init__()

    def read(self, reader, recid, reclen):
        return self.cls(None, None, None)


class WorkbookPrHandler(RecordHandler):
    cls = namedtuple('workbookPr', ['date1904', 'defaultThemeVersion'])

    def __init__(self):
        super(WorkbookPrHandler, self).__init__()

    def read(self, reader, recid, reclen):
        bits = reader.read_short() # TODO: This contains the 1904 flag
        reader.skip(2) # Not sure what this is, other flags probably
        theme = reader.read_int()
        reader.skip(4) # Also not sure, more flags?
        return self.cls(bits & 0x01 == 0x01, theme)


class SheetHandler(RecordHandler):
    cls = namedtuple('sheet', ['sheetId', 'rId', 'name'])

    def __init__(self):
        super(SheetHandler, self).__init__()

    def read(self, reader, recid, reclen):
        reader.skip(4)
        sheetid = reader.read_int()
        relid = reader.read_string()
        name = reader.read_string()
        return self.cls(sheetid, relid, name)


class DimensionHandler(RecordHandler):
    cls = namedtuple('dimension', ['r', 'c', 'h', 'w'])

    def __init__(self):
        super(DimensionHandler, self).__init__()

    def read(self, reader, recid, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1)


class ColumnHandler(RecordHandler):
    cls = namedtuple('col', ['c1', 'c2', 'width', 'style'])

    def __init__(self):
        super(ColumnHandler, self).__init__()

    def read(self, reader, recid, reclen):
        c1 = reader.read_int()
        c2 = reader.read_int()
        width = reader.read_int() / 256
        style = reader.read_int()
        return self.cls(c1, c2, width, style)


class RowHandler(RecordHandler):
    cls = namedtuple('row', ['r'])

    def __init__(self):
        super(RowHandler, self).__init__()

    def read(self, reader, recid, reclen):
        r = reader.read_int()
        return self.cls(r)


class CellHandler(RecordHandler):
    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def __init__(self):
        super(CellHandler, self).__init__()

    def read(self, reader, recid, reclen):
        col = reader.read_int()
        style = reader.read_int()

        val = None
        if recid == records.NUM:
            val = reader.read_rk()
        elif recid == records.BOOLERR:
            # TODO Map error values
            val = hex(reader.read_byte())
        elif recid == records.BOOL:
            val = reader.read_byte() != 0
        elif recid == records.FLOAT:
            val = reader.read_double()
        elif recid == records.STRING:
            val = reader.read_int()

        return self.cls(col, val, None, style)

class FormulaCellHandler(RecordHandler):
    cls = namedtuple('c', ['c', 'v', 'f', 'style'])

    def __init__(self):
        super(FormulaCellHandler, self).__init__()

    def read(self, reader, recid, reclen):
        col = reader.read_int()
        style = reader.read_int()

        val = None
        if recid == records.FORMULA_STRING:
            val = reader.read_string()
        elif recid == records.FORMULA_FLOAT:
            val = reader.read_double()
        elif recid == records.FORMULA_BOOL:
            val = reader.read_byte() != 0
        elif recid == records.FORMULA_BOOLERR:
            # TODO Map error values
            val = hex(reader.read_byte())

        formula = None
        # 0x0001 = Recalc always, 0x0002 = Calc on open, 0x0008 = Part of shared
        grbits = reader.read_short()
        sz = reader.read_int()
        if sz:
            buf = reader.read(sz)
            if len(buf) == sz:
                formula = Formula.parse(buf)

        return self.cls(col, val, formula, style)


class HyperlinkHandler(RecordHandler):
    cls = namedtuple('hyperlink', ['r', 'c', 'h', 'w', 'rId'])

    def __init__(self):
        super(HyperlinkHandler, self).__init__()

    def read(self, reader, recid, reclen):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_int()
        c2 = reader.read_int()
        rId = reader.read_string()
        return self.cls(r1, c1, r2 - r1 + 1, c2 - c1 + 1, rId)


class StringTableHandler(RecordHandler):
    cls = namedtuple('sst', ['count', 'uniqueCount'])

    def __init__(self):
        super(StringTableHandler, self).__init__()

    def read(self, reader, recid, reclen):
        count = reader.read_int()
        unique = reader.read_int()
        return self.cls(count, unique)


class StringInstanceHandler(RecordHandler):
    cls = namedtuple('si', ['t'])

    def __init__(self):
        super(StringInstanceHandler, self).__init__()

    def read(self, reader, recid, reclen):
        reader.skip(1)
        val = reader.read_string()
        return self.cls(val)
