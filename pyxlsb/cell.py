import sys

if sys.version_info > (3,):
    basestring = (str, bytes)
    long = int


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
    __slots__ = ('row', 'col', 'value', 'formula', 'style_id')

    def __init__(self, row, col, value=None, formula=None, style_id=None):
        self.row = row
        self.col = col
        self.value = value
        self.formula = formula
        self.style_id = style_id

    def __repr__(self):
        return 'Cell(row={}, col={}, value={}, formula={}, style_id={})' \
            .format(self.row, self.col, self.value, self.formula, self.style_id)

    @property
    def row_num(self):
        return self.row.num

    @property
    def string_value(self):
        if isinstance(self.value, basestring):
            return self.value

    @property
    def date_value(self):
        return self.row.sheet.workbook.convert_date(self.value)

    @property
    def numeric_value(self):
        if isinstance(self.value, (int, long, float)):
            return self.value

    @property
    def bool_value(self):
        if isinstance(self.value, bool):
            return self.value

    @property
    def is_date_formatted(self):
        fmt = self.row.sheet.workbook.styles._get_format(self.style_id)
        return fmt.is_date_format
