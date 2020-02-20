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

    @property
    def s(self):
        return self.style

class Cell(DeprecatedCellMixin):
    __slots__ = ('row_cls', 'row', 'col', 'value', 'value_conv', 'formula', 'style_num', 'style_fmt')

    def __init__(self, row, col, value=None, value_conv=None, formula=None, style_num=None, style_fmt=None):
        self.row_cls = row
        self.row = row.num
        self.col = col
        self.value = value
        self.value_conv = value_conv
        self.formula = formula
        self.style_num = style_num
        self.style_fmt = style_fmt

    def __repr__(self):
        return 'Cell(row_cls={}, row={}, col={}, value={}, value_conv={}, formula={}, style_num={}, style_fmt={})' \
            .format(self.row_cls, self.row, self.col, self.value, self.value_conv, self.formula, self.style_num, self.style_fmt)
