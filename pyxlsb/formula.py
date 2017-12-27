from .datareader import DataReader
from .tokenhandlers import *
from . import tokens

class Formula(object):
    def __init__(self, tokens=None):
        super(Formula, self).__init__()
        self._tokens = list(tokens) if tokens else []

    def __repr__(self):
        return 'Formula(%s)' % ([(hex(t[0]),) + t[1:] for t in self._tokens],)

    @classmethod
    def parse(cls, data):
        reader = data if hasattr(data, 'read') else DataReader(data)
        ptgs = []
        while True:
            tokenid = reader.read_byte()
            if tokenid is None:
                break
            baseid = ((tokenid | 0x20) if tokenid & 0x40 else tokenid) & 0x3F
            if tokens.sizes[baseid] > 0:
                token = (tokenid, reader.read(tokens.sizes[baseid]))
            else:
                token = (tokenid,)
            ptgs.append(token)
        return cls(ptgs)

