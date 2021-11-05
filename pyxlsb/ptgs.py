import sys

if sys.version_info > (3,):
    xrange = range


class BasePtg(object):
    def __repr__(self):
        args = ('{}={}'.format(str(k), repr(v)) for k, v in self.__dict__.items())
        return '{}({})'.format(self.__class__.__name__, ', '.join(args))

    def is_classified(self):
        return False

    def stringify(self, tokens, workbook):
        return '#PTG{}!'.format(self.ptg)

    @classmethod
    def read(cls, reader, ptg):
        return cls()

    def write(self, writer):
        # TODO Eventually, some day
        pass


class ClassifiedPtg(BasePtg):
    def __init__(self, ptg, *args, **kwargs):
        super(ClassifiedPtg, self).__init__(*args, **kwargs)
        self.ptg = ptg

    @property
    def base_ptg(self):
        return ((self.ptg | 0x20) if self.ptg & 0x40 == 0x40 else self.ptg) & 0x3F

    def is_classified(self):
        return True

    def is_reference(self):
        return self.ptg & 0x20 == 0x20 and self.ptg & 0x40 == 0

    def is_value(self):
        return self.ptg & 0x20 == 0 and self.ptg & 0x40 == 0x40

    def is_array(self):
        return self.ptg & 0x60 == 0x60

    @classmethod
    def read(cls, reader, ptg):
        return cls(ptg)


class UnknownPtg(BasePtg):
    ptg = 0xFF

    def __init__(self, ptg, *args, **kwargs):
        super(UnknownPtg, self).__init__(*args, **kwargs)
        self.ptg = ptg

    def stringify(self, tokens, workbook):
        return '#UNK{}!'.format(self.ptg)

    @classmethod
    def read(cls, reader, ptg):
        return cls(ptg)


# Unary operators


class UPlusPtg(BasePtg):
    ptg = 0x12

    def stringify(self, tokens, workbook):
        return '+' + tokens.pop().stringify(tokens, workbook)


class UMinusPtg(BasePtg):
    ptg = 0x13

    def stringify(self, tokens, workbook):
        return '-' + tokens.pop().stringify(tokens, workbook)


class PercentPtg(BasePtg):
    ptg = 0x14

    def stringify(self, tokens, workbook):
        return tokens.pop().stringify(tokens, workbook) + '%'


# Binary operators


class AddPtg(BasePtg):
    ptg = 0x03

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '+' + b


class SubstractPtg(BasePtg):
    ptg = 0x04

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '-' + b


class MultiplyPtg(BasePtg):
    ptg = 0x05

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '*' + b


class DividePtg(BasePtg):
    ptg = 0x06

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '/' + b


class PowerPtg(BasePtg):
    ptg = 0x07

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '^' + b


class ConcatPtg(BasePtg):
    ptg = 0x08

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '&' + b


class LessPtg(BasePtg):
    ptg = 0x09

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<' + b


class LessEqualPtg(BasePtg):
    ptg = 0x0A

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<=' + b


class EqualPtg(BasePtg):
    ptg = 0x0B

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '=' + b


class GreaterEqualPtg(BasePtg):
    ptg = 0x0C

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '>=' + b


class GreaterPtg(BasePtg):
    ptg = 0x0D

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '>' + b


class NotEqualPtg(BasePtg):
    ptg = 0x0E

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<>' + b


class IntersectionPtg(BasePtg):
    ptg = 0x0F

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ' ' + b


class UnionPtg(BasePtg):
    ptg = 0x10

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ',' + b


class RangePtg(BasePtg):
    ptg = 0x11

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ':' + b


# Operands


class MissArgPtg(BasePtg):
    ptg = 0x16

    def stringify(self, tokens, workbook):
        return ''


class StringPtg(BasePtg):
    ptg = 0x17

    def __init__(self, value, *args, **kwargs):
        super(StringPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return '"' + self.value.replace('"', '""') + '"'

    @classmethod
    def read(cls, reader, ptg):
        size = reader.read_short()
        value = reader.read_string(size=size)
        return cls(value)


class ErrorPtg(BasePtg):
    ptg = 0x1C

    def __init__(self, value, *args, **kwargs):
        super(ErrorPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        if self.value == 0x00:
            return '#NULL!'
        elif self.value == 0x07:
            return '#DIV/0!'
        elif self.value == 0x0F:
            return '#VALUE!'
        elif self.value == 0x17:
            return '#REF!'
        elif self.value == 0x1D:
            return '#NAME?'
        elif self.value == 0x24:
            return '#NUM!'
        elif self.value == 0x2A:
            return '#N/A'
        else:
            return '#ERR!'

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_byte()
        return cls(value)


class BooleanPtg(BasePtg):
    ptg = 0x1D

    def __init__(self, value, *args, **kwargs):
        super(BooleanPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return 'TRUE' if self.value else 'FALSE'

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_bool()
        return cls(value)


class IntegerPtg(BasePtg):
    ptg = 0x1E

    def __init__(self, value, *args, **kwargs):
        super(IntegerPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return str(self.value)

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_short()
        return cls(value)


class NumberPtg(BasePtg):
    ptg = 0x1F

    def __init__(self, value, *args, **kwargs):
        super(NumberPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return str(self.value)

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_double()
        return cls(value)


class ArrayPtg(ClassifiedPtg):
    ptg = 0x20

    def __init__(self, cols, rows, values, *args, **kwargs):
        super(ArrayPtg, self).__init__(*args, **kwargs)
        self.cols = cols
        self.rows = rows
        self.values = values

    @classmethod
    def read(cls, reader, ptg):
        cols = reader.read_byte()
        if cols == 0:
            cols = 256
        rows = reader.read_short()
        values = []
        for _ in xrange(cols * rows):
            flag = reader.read_byte()
            value = None
            if flag == 0x01:
                value = reader.read_double()
            elif flag == 0x02:
                size = reader.read_short()
                value = reader.read_string(size=size)
            values.append(value)
        return cls(cols, rows, values, ptg)


class NamePtg(ClassifiedPtg):
    ptg = 0x23

    def __init__(self, idx, reserved, *args, **kwargs):
        super(NamePtg, self).__init__(*args, **kwargs)
        self.idx = idx
        self._reserved = reserved

    # FIXME: We need to read names to stringify this

    @classmethod
    def read(cls, reader, ptg):
        idx = reader.read_short()
        res = reader.read(2)  # Reserved
        return cls(idx, res, ptg)


class RefPtg(ClassifiedPtg):
    ptg = 0x24

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens, workbook):
        r = ('R[{}]' if self.row_rel else 'R{}').format(self.row + 1)
        c = ('C[{}]' if self.col_rel else 'C{}').format(self.col + 1)
        return r + c

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(row, col & 0x3FFF, row_rel, col_rel, ptg)


class AreaPtg(ClassifiedPtg):
    ptg = 0x25

    def __init__(self, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel, last_col_rel, *args, **kwargs):
        super(AreaPtg, self).__init__(*args, **kwargs)
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        r1 = ('R[{}]' if self.first_row_rel else 'R{}').format(self.first_row + 1)
        c1 = ('C[{}]' if self.first_col_rel else 'C{}').format(self.first_col + 1)
        r2 = ('R[{}]' if self.last_row_rel else 'R{}').format(self.last_row + 1)
        c2 = ('C[{}]' if self.last_col_rel else 'C{}').format(self.last_col + 1)
        return r1 + c1 + ':' + r2 + c2

    @classmethod
    def read(cls, reader, ptg):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1rel = c1 & 0x8000 == 0x8000
        r2rel = c2 & 0x8000 == 0x8000
        c1rel = c1 & 0x4000 == 0x4000
        c2rel = c2 & 0x4000 == 0x4000
        return cls(r1, c1 & 0x3FFF, r2, c2 & 0x3FFF, r1rel, r2rel, c1rel, c2rel, ptg)


class MemAreaPtg(ClassifiedPtg):
    ptg = 0x26

    def __init__(self, reserved, rects, *args, **kwargs):
        super(ClassifiedPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self.rects = rects

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        rects = []
        if subex_len:
            rect_count = reader.read_short()
            for _ in xrange(rect_count):
                r1 = reader.read_int()
                r2 = reader.read_int()
                c1 = reader.read_short()
                c2 = reader.read_short()
                rects.append((r1, r2, c1, c2))
        return cls(res, rects, ptg)


class MemErrPtg(ClassifiedPtg):
    ptg = 0x27

    def __init__(self, reserved, subex, *args, **kwargs):
        super(MemErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self._subex = subex

    def stringify(self, tokens, workbook):
        return '#ERR!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(res, subex, ptg)


class RefErrPtg(ClassifiedPtg):
    ptg = 0x2A

    def __init__(self, reserved, *args, **kwargs):
        super(RefErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(6)  # Reserved
        return cls(res, ptg)


class AreaErrPtg(ClassifiedPtg):
    ptg = 0x2B

    def __init__(self, reserved, *args, **kwargs):
        super(AreaErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(12)  # Reserved
        return cls(res, ptg)


class RefNPtg(ClassifiedPtg):
    ptg = 0x2C

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefNPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens, workbook):
        r = ('R[{}]' if self.row_rel else 'R{}').format(self.row + 1)
        c = ('C[{}]' if self.col_rel else 'C{}').format(self.col + 1)
        return r + c

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(row, col & 0x3FFF, row_rel, col_rel, ptg)


class AreaNPtg(ClassifiedPtg):
    ptg = 0x2D

    def __init__(self, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel, last_col_rel, *args, **kwargs):
        super(AreaNPtg, self).__init__(*args, **kwargs)
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        r1 = ('R[{}]' if self.first_row_rel else 'R{}').format(self.first_row + 1)
        c1 = ('C[{}]' if self.first_col_rel else 'C{}').format(self.first_col + 1)
        r2 = ('R[{}]' if self.last_row_rel else 'R{}').format(self.last_row + 1)
        c2 = ('C[{}]' if self.last_col_rel else 'C{}').format(self.last_col + 1)
        return r1 + c1 + ':' + r2 + c2

    @classmethod
    def read(cls, reader, ptg):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1_rel = c1 & 0x8000 == 0x8000
        r2_rel = c2 & 0x8000 == 0x8000
        c1_rel = c1 & 0x4000 == 0x4000
        c2_rel = c2 & 0x4000 == 0x4000
        return cls(r1, r2, c1 & 0x3FFF, c2 & 0x3FFF, r1_rel, r2_rel, c1_rel, c2_rel, ptg)


class NameXPtg(ClassifiedPtg):
    ptg = 0x39

    def __init__(self, sheet_idx, name_idx, reserved, *args, **kwargs):
        super(NameXPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.name_idx = name_idx
        self._reserved = reserved

    # FIXME: We need to read names to stringify this

    @classmethod
    def read(cls, reader, ptg):
        sheet_idx = reader.read_short()
        name_idx = reader.read_short()
        res = reader.read(2)  # Reserved
        return cls(sheet_idx, name_idx, res, ptg)


class Ref3dPtg(ClassifiedPtg):
    ptg = 0x3A

    def __init__(self, sheet_idx, row, col, row_rel, col_rel, *args, **kwargs):
        super(Ref3dPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens, workbook):
        r = ('R[{}]' if self.row_rel else 'R{}').format(self.row + 1)
        c = ('C[{}]' if self.col_rel else 'C{}').format(self.col + 1)
        return workbook.sheets[self.sheet_idx] + '!' + r + c

    @classmethod
    def read(cls, reader, ptg):
        sheet_idx = reader.read_short()
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(sheet_idx, row, col & 0x3FFF, row_rel, col_rel, ptg)


class Area3dPtg(ClassifiedPtg):
    ptg = 0x3B

    def __init__(self, sheet_idx, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel, last_col_rel, *args, **kwargs):
        super(Area3dPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        r1 = ('R[{}]' if self.first_row_rel else 'R{}').format(self.first_row + 1)
        c1 = ('C[{}]' if self.first_col_rel else 'C{}').format(self.first_col + 1)
        r2 = ('R[{}]' if self.last_row_rel else 'R{}').format(self.last_row + 1)
        c2 = ('C[{}]' if self.last_col_rel else 'C{}').format(self.last_col + 1)
        return workbook.sheets[self.sheet_idx] + '!' + r1 + c1 + ':' + r2 + c2

    @classmethod
    def read(cls, reader, ptg):
        sheet_idx = reader.read_short()
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1_rel = c1 & 0x8000 == 0x8000
        r2_rel = c2 & 0x8000 == 0x8000
        c1_rel = c1 & 0x4000 == 0x4000
        c2_rel = c2 & 0x4000 == 0x4000
        return cls(sheet_idx, r1, r2, c1 & 0x3FFF, c2 & 0x3FFF, r1_rel, r2_rel, c1_rel, c2_rel, ptg)


class RefErr3dPtg(ClassifiedPtg):
    ptg = 0x3C

    def __init__(self, reserved, *args, **kwargs):
        super(RefErr3dPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(8)  # Reserved
        return cls(res, ptg)


class AreaErr3dPtg(ClassifiedPtg):
    ptg = 0x3D

    def __init__(self, reserved, *args, **kwargs):
        super(AreaErr3dPtg, self).__init__(*args, **kwargs)

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(14)  # Reserved
        return cls(res, ptg)


# Control


class ExpPtg(BasePtg):
    ptg = 0x01

    def __init__(self, row, col, *args, **kwargs):
        super(ExpPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col

    # FIXME: We need a workbook that supports direct cell referencing to stringify this

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        return cls(row, col)


class TablePtg(BasePtg):
    ptg = 0x02

    def __init__(self, row, col, *args, **kwargs):
        super(TablePtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col

    # FIXME: We need a workbook that supports tables to stringify this

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        return cls(row, col)


class ParenPtg(BasePtg):
    ptg = 0x15

    def stringify(self, tokens, workbook):
        return '(' + tokens.pop().stringify(tokens, workbook) + ')'


class AttrPtg(BasePtg):
    ptg = 0x19

    def __init__(self, flags, data, *args, **kwargs):
        super(AttrPtg, self).__init__(*args, **kwargs)
        self.flags = flags
        self.data = data

    @property
    def attr_semi(self):
        return self.flags & 0x01 == 0x01

    @property
    def attr_if(self):
        return self.flags & 0x02 == 0x02

    @property
    def attr_choose(self):
        return self.flags & 0x04 == 0x04

    @property
    def attr_goto(self):
        return self.flags & 0x08 == 0x08

    @property
    def attr_sum(self):
        return self.flags & 0x10 == 0x10

    @property
    def attr_baxcel(self):
        return self.flags & 0x20 == 0x20

    @property
    def attr_space(self):
        return self.flags & 0x40 == 0x40

    def stringify(self, tokens, workbook):
        spaces = ''
        if self.data & 0x00FF in [0x00, 0x06]:
            if self.attr_space:
                spaces = ' ' * (self.data >> 1)
        elif self.data & 0x00FF == 0x01:
            if self.attr_space:
                spaces = '\n' * (self.data >> 1)
        return spaces + tokens.pop().stringify(tokens, workbook)

    @classmethod
    def read(cls, reader, ptg):
        flags = reader.read_byte()
        data = reader.read_short()
        return cls(flags, data)


class MemNoMemPtg(ClassifiedPtg):
    ptg = 0x28

    def __init__(self, reserved, subex, *args, **kwargs):
        super(MemNoMemPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(res, subex, ptg)


class MemFuncPtg(ClassifiedPtg):
    ptg = 0x29

    def __init__(self, subex, *args, **kwargs):
        super(MemFuncPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(subex, ptg)


class MemAreaNPtg(ClassifiedPtg):
    ptg = 0x2E

    def __init__(self, subex, *args, **kwargs):
        super(MemAreaNPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.skip(subex_len)
        return cls(subex, ptg)


class MemNoMemNPtg(ClassifiedPtg):
    ptg = 0x2F

    def __init__(self, subex, *args, **kwargs):
        super(MemNoMemNPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(subex, ptg)

# Func operators


class FuncPtg(ClassifiedPtg):
    ptg = 0x21

    def __init__(self, idx, *args, **kwargs):
        super(FuncPtg, self).__init__(*args, **kwargs)
        self.idx = idx

    # FIXME: We need a workbook to stringify this (most likely)

    @classmethod
    def read(cls, reader, ptg):
        idx = reader.read_short()
        return cls(idx, ptg)


class FuncVarPtg(ClassifiedPtg):
    ptg = 0x22

    def __init__(self, idx, argc, prompt, ce, *args, **kwargs):
        super(FuncVarPtg, self).__init__(*args, **kwargs)
        self.idx = idx
        self.argc = argc
        self.prompt = prompt
        self.ce = ce

    def stringify(self, tokens, workbook):
        args = []
        for _ in xrange(self.argc):
            arg = tokens.pop().stringify(tokens, workbook)
            args.append(arg)
        return 'VAR_FUNC_{}({})'.format(self.idx, ', '.join(args))

    @classmethod
    def read(cls, reader, ptg):
        argc = reader.read_byte()
        idx = reader.read_short()
        return cls(idx & 0x7FFF, argc & 0x7F, argc & 0x80 == 0x80, idx & 0x8000 == 0x8000, ptg)
