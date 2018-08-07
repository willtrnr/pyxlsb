import re


class CellReference(object):
    _cell_ref_re = re.compile(r'^(\$)?([A-Z]+)(\$)?([0-9]+)$', re.I)

    def __init__(self, row, col, row_rel=False, col_rel=False):
        super(CellReference, self).__init__()
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def __eq__(self, other):
        return self.row == other.row and self.col == other.col and self.row_rel == other.row_rel and self.col_rel == other.col_rel

    def __repr__(self):
        return 'CellReference(row={}, col={}, row_rel={}, col_rel={})'.format(self.row, self.col, self.row_rel, self.col_rel)

    def __str__(self):
        c = ('' if self.col_rel else '$') + CellReference.index_to_col(self.col)
        r = ('' if self.row_rel else '$') + str(self.row + 1)
        return c + r

    @classmethod
    def parse(cls, value):
        match = cls._cell_ref_re.match(value)
        if match:
            col_rel = match.group(1) != '$'
            col = cls.col_to_index(match.group(2))
            row_rel = match.group(3) != '$'
            row = int(match.group(4)) - 1
            return cls(row, col, row_rel, col_rel)

    @staticmethod
    def col_to_index(col):
        idx = 0
        for c in col.upper():
            c = ord(c)
            if c >= ord('A') and c <= ord('Z'):
                idx = (idx * 26) + (c - ord('A') + 1)
            else:
                raise ValueError('invalid ALPHA-26 digit')
        return idx - 1

    @staticmethod
    def index_to_col(idx):
        idx += 1
        col = ''
        while idx > 0:
            digit = idx % 26
            if digit == 0:
                digit = 26
            idx = (idx - digit) // 26
            col = chr(ord('A') + digit - 1) + col
        return col
