import sys
from .cell import Cell

if sys.version_info > (3,):
    xrange = range


class Row(object):
    __slots__ = ('sheet', 'num', '_cols', '_cells')

    def __init__(self, sheet, num):
        self.sheet = sheet
        self.num = num
        self._cols = 0
        self._cells = dict()

    def __len__(self):
        return self._cols

    def __getitem__(self, key):
        if key in self._cells:
            return self._cells[key]
        else:
            return Cell(self, key, value=None)

    def __iter__(self):
        for i in xrange(self.__len__()):
            yield self.__getitem__(i)

    def _add_cell(self, col, *args, **kwargs):
        c = Cell(self, col, *args, **kwargs)
        self._cells[col] = c
        if col > self._cols:
            self._cols = col
        return c
