from datetime import datetime, timedelta
#from numpy import datetime64

def from_excel_ordinal_to_datetime(value, to_date=False, epoch=datetime(1900, 1, 1), cache={}):
    """
    https://stackoverflow.com/questions/29387137/how-to-convert-a-given-ordinal-number-from-excel-to-a-date
    epoch=datetime(1900, 1, 1) in function header for once initialization
    Adapted from above, thanks to @Martijn Pieters 
    """
    if not isinstance(value, (int, float)):
        return None
    if value in cache:
        return cache[value]

    if value > 60:
        value -= 1  # Excel leap year bug, 1900 is not a leap year!

    if to_date:
        return epoch + timedelta(days=int(value) - 1) # epoch is day 1
    else:
        inDays = int(value)
        inMicroSecs = int(round((value - inDays) * 86400000000.0))
        return epoch + timedelta(days=inDays - 1, microseconds=inMicroSecs) # epoch is day 1

def convert_value(value, dtype):
    if dtype == "datetime":
        #return datetime64(from_excel_ordinal_to_datetime(value))
        return from_excel_ordinal_to_datetime(value)
    elif dtype == "string":
        return str(value)
    else:
        return value

def detect_dtype(fmtCode):
    for char in ["d", "m", "y", "h", "m", "s"]:
        if char in fmtCode:
            return "datetime"
    for char in ["0.00", "0,00", "#.##", "#,##"]:
        if char in fmtCode:
            return "float64"
    return ""
