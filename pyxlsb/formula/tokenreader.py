from .ptgs import *
from pyxlsb.datareader import DataReader

class TokenReader(object):
    default_ptg = UnknownPtg

    ptgs = {
        # Unary operators
        UnaryPlusPtg.ptg:  UnaryPlusPtg,
        UnaryMinusPtg.ptg: UnaryMinusPtg,
        PercentPtg.ptg:    PercentPtg,

        # Binary operators
        AddPtg.ptg:              AddPtg,
        SubstractPtg.ptg:        SubstractPtg,
        MultiplyPtg.ptg:         MultiplyPtg,
        DividePtg.ptg:           DividePtg,
        PowerPtg.ptg:            PowerPtg,
        ConcatPtg.ptg:           ConcatPtg,
        LessThanPtg.ptg:         LessThanPtg,
        LessThanEqualPtg.ptg:    LessThanEqualPtg,
        EqualPtg.ptg:            EqualPtg,
        GreaterThanEqualPtg.ptg: GreaterThanEqualPtg,
        GreaterThanPtg.ptg:      GreaterThanPtg,
        NotEqualPtg.ptg:         NotEqualPtg,
        IntersectionPtg.ptg:     IntersectionPtg,
        UnionPtg.ptg:            UnknownPtg,
        RangePtg.ptg:            RangePtg,

        # Operands
        MissingArgPtg.ptg:       MissingArgPtg,
        StringPtg.ptg:           StringPtg,
        ErrorPtg.ptg:            ErrorPtg,
        BooleanPtg.ptg:          BooleanPtg,
        IntegerPtg.ptg:          IntegerPtg,
        NumberPtg.ptg:           NumberPtg,
        ArrayPtg.ptg:            ArrayPtg,
        NamePtg.ptg:             NamePtg,
        ReferencePtg.ptg:        ReferencePtg,
        AreaPtg.ptg:             AreaPtg,
        MemAreaPtg.ptg:          MemAreaPtg,
        MemErrorPtg.ptg:         MemErrorPtg,
        ReferenceErrorPtg.ptg:   ReferenceErrorPtg,
        AreaErrorPtg.ptg:        AreaErrorPtg,
        ReferenceNPtg.ptg:       ReferenceNPtg,
        AreaNPtg.ptg:            AreaNPtg,
        NameXPtg.ptg:            NameXPtg,
        Reference3dPtg.ptg:      Reference3dPtg,
        Area3dPtg.ptg:           Area3dPtg,
        ReferenceError3dPtg.ptg: ReferenceError3dPtg,
        AreaError3dPtg.ptg:      AreaError3dPtg,

        # Control
        ExpPtg.ptg:       ExpPtg,
        TblPtg.ptg:       TblPtg,
        ParenPtg.ptg:     ParenPtg,
        AttrPtg.ptg:      AttrPtg,
        MemNoMemPtg.ptg:  MemNoMemPtg,
        MemFuncPtg.ptg:   MemFuncPtg,
        MemAreaNPtg.ptg:  MemAreaNPtg,
        MemNoMemNPtg.ptg: MemNoMemNPtg,

        # Function operators
        FuncPtg.ptg: FuncPtg,
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
