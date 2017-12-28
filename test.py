from pyxlsb import open_workbook

with open_workbook('Test.xlsb', _debug=True) as wb:
    for name in wb.sheets:
        with wb.get_sheet(name) as sheet:
            for row in sheet:
                pass
