class CellDeprecated(object):
    __slots__ = ()

    @property
    def r(self):
        return self.row.num

    @property
    def c(self):
        return self.col

    @property
    def v(self):
        return self.value

    @property
    def f(self):
        return self.formula

class Cell(CellDeprecated):
    __slots__ = ('row', 'col', 'value', 'formula')

    def __init__(self, row, col, value, formula=None):
        self.row = row
        self.col = col
        self.value = value
        self.formula = formula

    def __repr__(self):
        return 'Cell(r={}, c={}, v={}, f={})'.format(self.r, self.c, self.value, self._formula)
