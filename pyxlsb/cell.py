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
    __slots__ = ('row', 'col', 'value', 'formula', 'style_num')

    def __init__(self, row, col, value=None, formula=None, style_num=None):
        self.row = row
        self.col = col
        self.value = value
        self.formula = formula
        self.style_num = style_num

    def __repr__(self):
        return 'Cell(row={}, row_num={}, col={}, value={}, formula={}, style_num={})' \
            .format(self.row, self.row_num, self.col, self.value, self.formula, self.style_num)

    @property
    def row_num(self):
        return self.row.num

    @property
    def value_conv(self):
        dtype = self.row.sheet.workbook.styles.get_dtype(self.style_num)
        if dtype == "datetime":
            return self.row.sheet.workbook.convert_date(self.value)
        elif dtype == "string":
            return str(self.value)
        else:
            return self.value

    @property
    def format(self):
        return self.row.sheet.workbook.styles.get_format(self.style_num)
