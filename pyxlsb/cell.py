import sys

if sys.version_info > (3,):
    basestring = (str, bytes)
    long = int


class DeprecatedCellMixin(object):
    """Deprecated Cell properties to preserve source compatibility with the 1.0.x releases."""

    __slots__ = ()

    @property
    def r(self):
        """The row number of this cell.

        .. deprecated:: 1.1.0
            Use the ``row_num`` property instead.
        """
        return self.row.num

    @property
    def c(self):
        """The column number of this cell.

        .. deprecated:: 1.1.0
            Use the ``col`` property instead.
        """
        return self.col

    @property
    def v(self):
        """The value of this cell.

        .. deprecated:: 1.1.0
            Use the ``value`` or the typed ``*_value`` properties instead.
        """
        return self.value

    @property
    def f(self):
        """The formula of this cell.

        .. deprecated:: 1.1.0
            Use the ``formula`` property instead.
        """
        return self.formula

class Cell(DeprecatedCellMixin):
    """A cell in a worksheet.

    Attributes:
        row (Row): The containing row.
        col (int): The column index for this cell.
        value (mixed): The cell value.
        formula (bytes): The formula PTG bytes.
        style_id (int): The style index in the style table.
    """

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
        """The row number of this cell."""
        return self.row.num

    @property
    def string_value(self):
        """The string value of this cell or None if not a string."""
        if isinstance(self.value, basestring):
            return self.value

    @property
    def numeric_value(self):
        """The numeric value of this cell or None if not a number."""
        if isinstance(self.value, (int, long, float)):
            return self.value

    @property
    def bool_value(self):
        """The boolean value of this cell or None if not a boolean."""
        if isinstance(self.value, bool):
            return self.value

    @property
    def date_value(self):
        """The date value of this cell or None if not a numeric cell."""
        return self.row.sheet.workbook.convert_date(self.value)

    @property
    def is_date_formatted(self):
        """If this cell is formatted using a date-like format code."""
        fmt = self.row.sheet.workbook.styles._get_format(self.style_id)
        return fmt.is_date_format
