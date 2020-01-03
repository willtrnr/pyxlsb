class DeprecatedCellMixin(object):
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

class Cell(DeprecatedCellMixin):
    __slots__ = ('row', 'col', 'value', 'formula')

    def __init__(self, row, col, value=None, formula=None):
        self.row = row
        self.col = col
        self.value = value
        self.formula = formula

    def __repr__(self):
        return 'Cell(row={}, col={}, value={}, formula={})' \
            .format(self.row, self.col, self.value, self.formula)
