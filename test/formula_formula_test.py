import unittest
from pyxlsb.formula import Formula
from pyxlsb.formula.ptgs import *

class FormulaTestCase(unittest.TestCase):
    def test_stringify(self):
        self.assertEqual(Formula([StringPtg('A')]).stringify(), '"A"')
        self.assertEqual(Formula([IntegerPtg(1), IntegerPtg(2), AddPtg()]).stringify(), '1+2')
