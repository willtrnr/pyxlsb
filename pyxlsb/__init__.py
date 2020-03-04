import warnings
from datetime import datetime, timedelta
from .workbook import Workbook
from .xlsbpackage import XlsbPackage

__version__ = '1.1.0'


def open_workbook(name, *args, **kwargs):
    """Opens the given workbook file path.

    Args:
        name (str): The name of the XLSB file to open.

    Returns:
        Workbook: The workbook instance for the given file name.

    Examples:
        This is typically the entrypoint to start working with an XLSB file:

        >>> from pyxlsb import open_workbook
        >>> with open_workbook('test_files/test.xlsb') as wb:
        ...     print(wb.sheets)
        ...
        ['Test']
    """
    return Workbook(XlsbPackage(name), *args, **kwargs)


def convert_date(date):
    """Convert an Excel date to a ``datetime`` instance.

    Args:
        date (int or float): The numeric Excel date value to convert.

    Returns:
        datetime: The corresponding datetime instance or None if invalid.

    .. deprecated:: 1.1.0
        Will be removed in 1.2.0. Use ``convert_date`` on ``Workbook`` instances instead.
    """
    warnings.warn("convert_date was moved to the Workbook object", DeprecationWarning)
    if not isinstance(date, (int, float)):
        return None

    if int(date) == 0:
        return datetime(1900, 1, 1, 0, 0, 0) + timedelta(seconds=date * 24 * 60 * 60)
    elif date >= 61:
        # According to Lotus 1-2-3, Feb 29th 1900 is a real thing, therefore we have to remove one day after that date
        return datetime(1899, 12, 31, 0, 0, 0) + timedelta(days=int(date) - 1, seconds=int((date % 1) * 24 * 60 * 60))
    else:
        # Feb 29th 1900 will show up as Mar 1st 1900 because Python won't handle that date
        return datetime(1899, 12, 31, 0, 0, 0) + timedelta(days=int(date), seconds=int((date % 1) * 24 * 60 * 60))
