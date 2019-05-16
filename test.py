import sys
import time
from pyxlsb import open_workbook

a = time.time()
print('Opening workbook... ', end='', flush=True)
with open_workbook(sys.argv[1]) as wb:
    d = time.time() - a
    print('Done! ({} seconds)'.format(d))
    for s in wb.sheets:
        print('Reading sheet {}... '.format(s), end='', flush=True)
        a = time.time()
        with wb.get_sheet(s) as sheet:
            for row in sheet:
                pass
        d = time.time() - a
        print('Done! ({} seconds)'.format(d))
