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
        return '#TODO{}!'.format(self.ptg)

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

class UnaryPlusPtg(BasePtg):
    ptg = 0x12

    def stringify(self, tokens):
        return '+' + tokens.pop().stringify(tokens)

class UnaryMinusPtg(BasePtg):
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

class LessThanPtg(BasePtg):
    ptg = 0x09

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '<' + b

class LessThanEqualPtg(BasePtg):
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

class GreaterThanEqualPtg(BasePtg):
    ptg = 0x0C

    def stringify(self, tokens):
        b = tokens.pop().stringify(tokens)
        a = tokens.pop().stringify(tokens)
        return a + '>=' + b

class GreaterThanPtg(BasePtg):
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

class MissingArgPtg(BasePtg):
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

    def __init__(self, index, *args, **kwargs):
        super(NamePtg, self).__init__(*args, **kwargs)
        self.index = index

    # FIXME: We need a workbook to stringify this

    @classmethod
    def parse(cls, reader, ptg):
        idx = reader.read_short()
        reader.skip(2) # Reserved
        return cls(idx, ptg)

class ReferencePtg(ClassifiedPtg):
    ptg = 0x24

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(ReferencePtg, self).__init__(*args, **kwargs)
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

class MemErrorPtg(ClassifiedPtg):
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

class ReferenceErrorPtg(ClassifiedPtg):
    ptg = 0x2A

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(6) # Reserved
        return cls(ptg)

class AreaErrorPtg(ClassifiedPtg):
    ptg = 0x2B

    def stringify(self, tokens):
        return '#REF!'

    @classmethod
    def parse(cls, reader, ptg):
        reader.skip(12)
        return cls(ptg)

class ReferenceNPtg(ClassifiedPtg):
    ptg = 0x2C

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(ReferenceNPtg, self).__init__(*args, **kwargs)
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
