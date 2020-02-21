from .conv import convert_value

class DeprecatedCellMixin(object):
    __slots__ = ()

    @property
    def r(self):
        return self.row_cls.num

    @property
    def c(self):
        return self.col

    @property
    def v(self):
        return self.value

    @property
    def value_conv(self):
        return convert_value(self.value, self.number_format)

    @property
    def number_format(self):
        return self.row_cls.sheet.workbook.styles.get_number_format(self.style_num)

    @property
    def f(self):
        return self.formula

    @property
    def s(self):
        return self.style

class Cell(DeprecatedCellMixin):
    __slots__ = ('row_cls', 'row', 'col', 'value', 'formula', 'style_num')

    def __init__(self, row_cls, col, value=None, formula=None, style_num=None):
        self.row_cls = row_cls
        self.row = row_cls.num
        self.col = col
        self.value = value
        self.formula = formula
        self.style_num = style_num

    def __repr__(self):
        return 'Cell(row_cls={}, row={}, col={}, value={}, formula={}, style_num={})' \
            .format(self.row_cls, self.row, self.col, self.value, self.formula, self.style_num)
