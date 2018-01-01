import re

class CellReference(object):
    _cell_ref_re = r'^(\$)?([A-Z]+)(\$)?([0-9]+)$'

    def __init__(self, row, col, row_rel=False, col_rel=False):
        super(CellReference, self).__init__()
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    @classmethod
    def parse(cls, value):
        match = cls._cell_ref_re.match(value)
        if match:
            col_rel = match.group(1) == '$'
            col = cls.col_to_index(match.group(2))
            row_rel = match.group(3) == '$'
            row = int(match.group(4))
            return cls(row, col, row_rel, col_rel)

    @staticmethod
    def col_to_index(col):
        idx = 0
        for c in col.upper():
            c = ord(c)
            if c >= ord('A') and c <= ord('Z'):
                idx *= 26
                idx += c - ord('A') + 1
        return idx - 1
