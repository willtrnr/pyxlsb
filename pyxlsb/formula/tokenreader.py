from .ptgs import *
from pyxlsb.datareader import DataReader

class TokenReader(object):
    default_ptg = UnknownPtg

    ptgs = {
        # Unary
        UnaryPlusPtg.ptg:  UnaryPlusPtg,
        UnaryMinusPtg.ptg: UnaryMinusPtg,
        PercentPtg.ptg:    PercentPtg,

        # Binary
        AddPtg.ptg:       AddPtg,
        SubstractPtg.ptg: SubstractPtg,
        MultiplyPtg.ptg:  MultiplyPtg,
        DividePtg.ptg:    DividePtg
    }

    def __init__(self, fp):
        self._fp = fp

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def next(self):
        reader = DataReader(self._fp)
        ptg = reader.read_byte()
        if ptg is None:
            raise StopIteration
        base = ((ptg | 0x20) if ptg & 0x40 == 0x40 else ptg) & 0x3F
        return (self.ptgs.get(base) or self.default_ptg).parse(reader, ptg)

    def close(self):
        self._fp.close()
