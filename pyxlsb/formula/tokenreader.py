from .ptgs import *
from pyxlsb.datareader import DataReader

class TokenReader(object):
    default_ptg = UnknownPtg

    ptgs = {
        # Unary operators
        UPlusPtg.ptg:   UPlusPtg,
        UMinusPtg.ptg:  UMinusPtg,
        PercentPtg.ptg: PercentPtg,

        # Binary operators
        AddPtg.ptg:          AddPtg,
        SubstractPtg.ptg:    SubstractPtg,
        MultiplyPtg.ptg:     MultiplyPtg,
        DividePtg.ptg:       DividePtg,
        PowerPtg.ptg:        PowerPtg,
        ConcatPtg.ptg:       ConcatPtg,
        LessPtg.ptg:         LessPtg,
        LessEqualPtg.ptg:    LessEqualPtg,
        EqualPtg.ptg:        EqualPtg,
        GreaterEqualPtg.ptg: GreaterEqualPtg,
        GreaterPtg.ptg:      GreaterPtg,
        NotEqualPtg.ptg:     NotEqualPtg,
        IntersectionPtg.ptg: IntersectionPtg,
        UnionPtg.ptg:        UnionPtg,
        RangePtg.ptg:        RangePtg,

        # Operands
        MissArgPtg.ptg:   MissArgPtg,
        StringPtg.ptg:    StringPtg,
        ErrorPtg.ptg:     ErrorPtg,
        BooleanPtg.ptg:   BooleanPtg,
        IntegerPtg.ptg:   IntegerPtg,
        NumberPtg.ptg:    NumberPtg,
        ArrayPtg.ptg:     ArrayPtg,
        NamePtg.ptg:      NamePtg,
        RefPtg.ptg:       RefPtg,
        AreaPtg.ptg:      AreaPtg,
        MemAreaPtg.ptg:   MemAreaPtg,
        MemErrPtg.ptg:    MemErrPtg,
        RefErrPtg.ptg:    RefErrPtg,
        AreaErrPtg.ptg:   AreaErrPtg,
        RefNPtg.ptg:      RefNPtg,
        AreaNPtg.ptg:     AreaNPtg,
        NameXPtg.ptg:     NameXPtg,
        Ref3dPtg.ptg:     Ref3dPtg,
        Area3dPtg.ptg:    Area3dPtg,
        RefErr3dPtg.ptg:  RefErr3dPtg,
        AreaErr3dPtg.ptg: AreaErr3dPtg,

        # Control
        ExpPtg.ptg:       ExpPtg,
        TablePtg.ptg:     TablePtg,
        ParenPtg.ptg:     ParenPtg,
        AttrPtg.ptg:      AttrPtg,
        MemNoMemPtg.ptg:  MemNoMemPtg,
        MemFuncPtg.ptg:   MemFuncPtg,
        MemAreaNPtg.ptg:  MemAreaNPtg,
        MemNoMemNPtg.ptg: MemNoMemNPtg,

        # Func operators
        FuncPtg.ptg:    FuncPtg,
        FuncVarPtg.ptg: FuncVarPtg
    }

    def __init__(self, fp):
        self._reader = DataReader(fp)

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def next(self):
        ptg = self._reader.read_byte()
        if ptg is None:
            raise StopIteration
        base = ((ptg | 0x20) if ptg & 0x40 == 0x40 else ptg) & 0x3F
        return (self.ptgs.get(base) or self.default_ptg).parse(self._reader, ptg)
