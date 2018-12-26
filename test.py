import sys
from pyxlsb import open_workbook

with open_workbook(sys.argv[1]) as wb:
    for i in range(len(wb.sheets)):
        with wb.get_sheet(i) as sheet:
            for row in sheet:
                pass
