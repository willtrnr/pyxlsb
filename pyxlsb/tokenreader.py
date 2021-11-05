from . import ptgs

try:
    from .cdatareader import DataReader
except ImportError:
    from .datareader import DataReader


class TokenReader(object):
    _default_ptg = ptgs.UnknownPtg

    _ptgs = {
        # Unary operators
        ptgs.UPlusPtg.ptg:   ptgs.UPlusPtg,
        ptgs.UMinusPtg.ptg:  ptgs.UMinusPtg,
        ptgs.PercentPtg.ptg: ptgs.PercentPtg,

        # Binary operators
        ptgs.AddPtg.ptg:          ptgs.AddPtg,
        ptgs.SubstractPtg.ptg:    ptgs.SubstractPtg,
        ptgs.MultiplyPtg.ptg:     ptgs.MultiplyPtg,
        ptgs.DividePtg.ptg:       ptgs.DividePtg,
        ptgs.PowerPtg.ptg:        ptgs.PowerPtg,
        ptgs.ConcatPtg.ptg:       ptgs.ConcatPtg,
        ptgs.LessPtg.ptg:         ptgs.LessPtg,
        ptgs.LessEqualPtg.ptg:    ptgs.LessEqualPtg,
        ptgs.EqualPtg.ptg:        ptgs.EqualPtg,
        ptgs.GreaterEqualPtg.ptg: ptgs.GreaterEqualPtg,
        ptgs.GreaterPtg.ptg:      ptgs.GreaterPtg,
        ptgs.NotEqualPtg.ptg:     ptgs.NotEqualPtg,
        ptgs.IntersectionPtg.ptg: ptgs.IntersectionPtg,
        ptgs.UnionPtg.ptg:        ptgs.UnionPtg,
        ptgs.RangePtg.ptg:        ptgs.RangePtg,

        # Operands
        ptgs.MissArgPtg.ptg:   ptgs.MissArgPtg,
        ptgs.StringPtg.ptg:    ptgs.StringPtg,
        ptgs.ErrorPtg.ptg:     ptgs.ErrorPtg,
        ptgs.BooleanPtg.ptg:   ptgs.BooleanPtg,
        ptgs.IntegerPtg.ptg:   ptgs.IntegerPtg,
        ptgs.NumberPtg.ptg:    ptgs.NumberPtg,
        ptgs.ArrayPtg.ptg:     ptgs.ArrayPtg,
        ptgs.NamePtg.ptg:      ptgs.NamePtg,
        ptgs.RefPtg.ptg:       ptgs.RefPtg,
        ptgs.AreaPtg.ptg:      ptgs.AreaPtg,
        ptgs.MemAreaPtg.ptg:   ptgs.MemAreaPtg,
        ptgs.MemErrPtg.ptg:    ptgs.MemErrPtg,
        ptgs.RefErrPtg.ptg:    ptgs.RefErrPtg,
        ptgs.AreaErrPtg.ptg:   ptgs.AreaErrPtg,
        ptgs.RefNPtg.ptg:      ptgs.RefNPtg,
        ptgs.AreaNPtg.ptg:     ptgs.AreaNPtg,
        ptgs.NameXPtg.ptg:     ptgs.NameXPtg,
        ptgs.Ref3dPtg.ptg:     ptgs.Ref3dPtg,
        ptgs.Area3dPtg.ptg:    ptgs.Area3dPtg,
        ptgs.RefErr3dPtg.ptg:  ptgs.RefErr3dPtg,
        ptgs.AreaErr3dPtg.ptg: ptgs.AreaErr3dPtg,

        # Control
        ptgs.ExpPtg.ptg:       ptgs.ExpPtg,
        ptgs.TablePtg.ptg:     ptgs.TablePtg,
        ptgs.ParenPtg.ptg:     ptgs.ParenPtg,
        ptgs.AttrPtg.ptg:      ptgs.AttrPtg,
        ptgs.MemNoMemPtg.ptg:  ptgs.MemNoMemPtg,
        ptgs.MemFuncPtg.ptg:   ptgs.MemFuncPtg,
        ptgs.MemAreaNPtg.ptg:  ptgs.MemAreaNPtg,
        ptgs.MemNoMemNPtg.ptg: ptgs.MemNoMemNPtg,

        # Func operators
        ptgs.FuncPtg.ptg:    ptgs.FuncPtg,
        ptgs.FuncVarPtg.ptg: ptgs.FuncVarPtg
    }

    def __init__(self, buf):
        self._reader = DataReader(buf)

    def __iter__(self):
        return self

    def __next__(self):
        return self.next()

    def next(self):
        ptg = self._reader.read_byte()
        if ptg is None:
            raise StopIteration
        base = ((ptg | 0x20) if ptg & 0x40 == 0x40 else ptg) & 0x3F
        return self._ptgs.get(base, self._default_ptg).read(self._reader, ptg)
