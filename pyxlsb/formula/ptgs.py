import sys

if sys.version_info > (3,):
    xrange = range

class BasePtg(object):
    def __repr__(self):
        args = list('{}={}'.format(str(k), repr(v)) for k, v in self.__dict__.items())
        return '{}({})'.format(self.__class__.__name__, ', '.join(args))

    def is_classified(self):
        return False

    def stringify(self, tokens):
        return '#PTG{}!'.format(self.ptg)

    @classmethod
    def parse(cls, reader, ptg):
        return cls()


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
    def parse(cls, reader, ptg):
        return cls(ptg)


class UnknownPtg(BasePtg):
    ptg = 0x00

    def __init__(self, ptg, *args, **kwargs):
        super(UnknownPtg, self).__init__(*args, **kwargs)
        self.ptg = ptg

    def stringify(self, tokens):
        return '#UNK{}!'.format(self.ptg)

    @classmethod
    def parse(cls, reader, ptg):
        return cls(ptg)

# Unary operators

class UPlusPtg(BasePtg):
    ptg = 0x12

    def stringify(self, tokens):
        return '+' + tokens.pop().stringify(tokens)

class UMinusPtg(BasePtg):
    ptg = 0x13

    def stringify(self, tokens):
        return '-' + tokens.pop().stringify(tokens)

class PercentPtg(BasePtg):
    ptg = 0x14

    def stringify(self, tokens):
        return tokens.pop().stringify(tokens) + '%'

# Binary operators

class AddPtg(BasePtg):
    ptg = 0x03

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '+' + b

class SubstractPtg(BasePtg):
    ptg = 0x04

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '-' + b

class MultiplyPtg(BasePtg):
    ptg = 0x05

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '*' + b

class DividePtg(BasePtg):
    ptg = 0x06

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '/' + b

class PowerPtg(BasePtg):
    ptg = 0x07

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '^' + b

class ConcatPtg(BasePtg):
    ptg = 0x08

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '&' + b

class LessPtg(BasePtg):
    ptg = 0x09

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '<' + b

class LessEqualPtg(BasePtg):
    ptg = 0x0A

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '<=' + b

class EqualPtg(BasePtg):
    ptg = 0x0B

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '=' + b

class GreaterEqualPtg(BasePtg):
    ptg = 0x0C

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '>=' + b

class GreaterPtg(BasePtg):
    ptg = 0x0D

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '>' + b

class NotEqualPtg(BasePtg):
    ptg = 0x0E

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '<>' + b

class IntersectionPtg(BasePtg):
    ptg = 0x0F

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + ' ' + b

class UnionPtg(BasePtg):
    ptg = 0x10

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + ',' + b

class RangePtg(BasePtg):
    ptg = 0x11

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + ':' + b

# Operands

class MissArgPtg(BasePtg):
    ptg = 0x16

    def stringify(self, tokens):
        return ''

class StringPtg(BasePtg):
    ptg = 0x17

    def __init__(self, value, *args, **kwargs):
        super(StringPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens):
        return '"{}"'.format(self.value.replace('"', '""'))

    @classmethod
    def parse(cls, reader, ptg):
        size = reader.read_short()
        value = reader.read_string(size=size)
        return cls(value)

class ErrorPtg(BasePtg):
    ptg = 0x1C

    def __init__(self, value, *args, **kwargs):
        super(ErrorPtg, self).__init__(*args, **kwargs)
        self.value = value

    @classmethod
    def parse(cls, reader, ptg):
        value = reader.read_byte()
        return cls(hex(value))

class BooleanPtg(BasePtg):
    ptg = 0x1D

    def __init__(self, value, *args, **kwargs):
        super(BooleanPtg, self).__init__
        self.value = value

    def stringify(self, tokens):
        return 'TRUE' if self.value else 'FALSE'

    @classmethod
    def parse(cls, reader, ptg):
        value = reader.read_bool()
        return cls(value)

class IntegerPtg(BasePtg):
    ptg = 0x1E

    def __init__(self, value, *args, **kwargs):
        super(IntegerPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens):
        return str(self.value)

    @classmethod
    def parse(cls, reader, ptg):
        value = reader.read_short()
        return cls(value)

class NumberPtg(BasePtg):
    ptg = 0x1F

    def __init__(self, value, *args, **kwargs):
        super(NumberPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens):
        return str(self.value)

    @classmethod
    def parse(cls, reader, ptg):
        value = reader.read_double()
        return cls(value)

class ArrayPtg(ClassifiedPtg):
    ptg = 0x20

    def __init__(self, cols, rows, values, *args, **kwargs):
        super(ArrayPtg, self).__init__(*args, **kwargs)
        self.cols = cols
        self.rows = rows
        self.values = values

    def stringify(self, tokens):
        return '{TODO}'

    @classmethod
    def parse(cls, reader, ptg):
        cols = reader.read_byte()
        if cols == 0:
            cols = 256
        rows = reader.read_short()
        values = list()
        for i in xrange(cols * rows):
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

    def __init__(self, idx, *args, **kwargs):
        super(NamePtg, self).__init__(*args, **kwargs)
        self.idx = idx

    # FIXME: We need a workbook to stringify this

    @classmethod
    def parse(cls, reader, ptg):
        idx = reader.read_short()
        reader.skip(2) # Reserved
        return cls(idx, ptg)

class RefPtg(ClassifiedPtg):
    ptg = 0x24

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens):
        # FIXME: Use A1
        r = ('R[{}]' if self.row_rel else 'R{}').format(self.row + 1)
        c = ('C[{}]' if self.col_rel else 'C{}').format(self.col + 1)
        return r + c

    @classmethod
    def parse(cls, reader, ptg):
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

    def stringify(self, tokens):
        # FIXME: Use A1
        r1 = ('R[{}]' if self.first_row_rel else 'R{}').format(self.first_row + 1)
        c1 = ('C[{}]' if self.first_col_rel else 'C{}').format(self.first_col + 1)
        r1 = ('R[{}]' if self.last_row_rel else 'R{}').format(self.last_row + 1)
        c1 = ('C[{}]' if self.last_col_rel else 'C{}').format(self.last_col + 1)
        return r1 + c1 + ':' + r2 + c2

    @classmethod
    def parse(cls, reader, ptg):
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

    def __init__(self, rects, *args, **kwargs):
        super(ClassifiedPtg, self).__init__(*args, **kwargs)
        self.rects = rects

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(4) # Reserved
        subex_len = reader.read_short()
        rects = list()
        if subex_len:
            rect_count = reader.read_short()
            for i in xrange(rect_count):
                r1 = reader.read_int()
                r2 = reader.read_int()
                c1 = reader.read_short()
                c2 = reader.read_short()
                rects.append((r1, r2, c1, c2))
        return cls(rects, ptg)

class MemErrPtg(ClassifiedPtg):
    ptg = 0x27

    def stringify(self, tokens):
        return '#ERR!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(4) # Reserved
        subex_len = reader.read_short()
        # TODO: The "spec" is very vague about this
        reader.skip(subex_len)
        return cls(ptg)

class RefErrPtg(ClassifiedPtg):
    ptg = 0x2A

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(6) # Reserved
        return cls(ptg)

class AreaErrPtg(ClassifiedPtg):
    ptg = 0x2B

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(12) # Reserved
        return cls(ptg)

class RefNPtg(ClassifiedPtg):
    ptg = 0x2C

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefNPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens):
        # FIXME: Use A1
        r = ('R[{}]' if self.row_rel else 'R{}').format(self.row + 1)
        c = ('C[{}]' if self.col_rel else 'C{}').format(self.col + 1)
        return r + c

    @classmethod
    def parse(cls, reader, ptg):
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

    def stringify(self, tokens):
        # FIXME: Use A1
        r1 = ('R[{}]' if self.first_row_rel else 'R{}').format(self.first_row + 1)
        c1 = ('C[{}]' if self.first_col_rel else 'C{}').format(self.first_col + 1)
        r1 = ('R[{}]' if self.last_row_rel else 'R{}').format(self.last_row + 1)
        c1 = ('C[{}]' if self.last_col_rel else 'C{}').format(self.last_col + 1)
        return r1 + c1 + ':' + r2 + c2

    @classmethod
    def parse(cls, reader, ptg):
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

    def __init__(self, sheet_idx, name_idx, *args, **kwargs):
        super(NameXPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.name_idx = name_idx

    # FIXME: We need a workbook to stringify this

    @classmethod
    def parse(cls, reader, ptg):
        sheet_idx = reader.read_short()
        name_idx = reader.read_short()
        reader.skip(2) # Reserved
        return cls(sheet_idx, name_idx, ptg)

class Ref3dPtg(ClassifiedPtg):
    ptg = 0x3A

    def __init__(self, sheet_idx, row, col, row_rel, col_rel, *args, **kwargs):
        super(Ref3dPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    # FIXME: We need a workbook to stringify this

    @classmethod
    def parse(cls, reader, ptg):
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

    # FIXME: We need a workbook to stringify this

    @classmethod
    def parse(cls, reader, ptg):
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

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(8) # Reserved
        return cls(ptg)

class AreaErr3dPtg(ClassifiedPtg):
    ptg = 0x3D

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(14) # Reserved
        return cls(ptg)

# Control

class ExpPtg(BasePtg):
    ptg = 0x01

    def __init__(self, row, col, *args, **kwargs):
        super(ExpPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col

    # FIXME: We need a workbook to stringify this

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

    # FIXME: We need a workbook to stringify this

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        return cls(row, col)

class ParenPtg(BasePtg):
    ptg = 0x15

    def stringify(self, tokens):
        return '(' + tokens.pop().stringify(tokens) + ')'

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

    def stringify(self, tokens):
        # TODO: Handle space stuff
        return tokens.pop().stringify(tokens)

    @classmethod
    def parse(cls, reader, ptg):
        flags = reader.read_byte()
        data = reader.read_short()
        return cls(flags, data)

class MemNoMemPtg(ClassifiedPtg):
    ptg = 0x28

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(4) # Reserved
        subex_len = reader.read_short()
        # TODO The "spec" is really vague about this
        reader.skip(subex_len)
        return cls(ptg)

class MemFuncPtg(ClassifiedPtg):
    ptg = 0x29

    @classmethod
    def parse(cls, reader, ptg):
        subex_len = reader.read_short()
        # TODO The "spec" is also really vague about this
        reader.skip(subex_len)
        return cls(ptg)

class MemAreaNPtg(ClassifiedPtg):
    ptg = 0x2E

    @classmethod
    def parse(cls, reader, ptg):
        subex_len = reader.read_short()
        # TODO Same thing, really vague info
        reader.skip(subex_len)
        return cls(ptg)

class MemNoMemNPtg(ClassifiedPtg):
    ptg = 0x2F

    @classmethod
    def parse(cls, reader, ptg):
        subex_len = reader.read_short()
        # TODO Still no idea what this is
        reader.skip(subex_len)
        return cls(ptg)

# Func operators

class FuncPtg(ClassifiedPtg):
    ptg = 0x21

    def __init__(self, idx, *args, **kwargs):
        super(FuncPtg, self).__init__(*args, **kwargs)
        self.idx = idx

    # FIXME: We need a workbook to stringify this (most likely)

    @classmethod
    def parse(cls, reader, ptg):
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

    def stringify(self, tokens):
        args = list()
        for i in xrange(self.argc):
            arg = tokens.pop().stringify(tokens)
            args.append(arg)
        return 'VAR_FUNC_{}({})'.format(self.idx, ', '.join(args))

    @classmethod
    def parse(cls, reader, ptg):
        argc = reader.read_byte()
        idx = reader.read_short()
        return cls(idx & 0x7FFF, argc & 0x7F, argc & 0x80 == 0x80, idx & 0x8000 == 0x8000, ptg)
