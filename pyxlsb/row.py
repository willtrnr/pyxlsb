class Row(object):
    __slots__ = ('sheet', 'num')

    def __init__(self, sheet, num):
        self.sheet = sheet
        self.num = num
