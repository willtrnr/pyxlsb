import unittest
from mock import Mock
from pyxlsb.formula import Formula
from pyxlsb.formula.ptgs import *

class FormulaTestCase(unittest.TestCase):
    def setUp(self):
        self.wb = Mock(sheets=['Sheet1', 'Sheet2'])

    def test_stringify(self):
        self.assertEqual(Formula([StringPtg('Hello "World"')]).stringify(self.wb), '"Hello ""World"""')
        self.assertEqual(Formula([IntegerPtg(1), IntegerPtg(2), AddPtg()]).stringify(self.wb), '1+2')
        self.assertEqual(Formula([Ref3dPtg(1, 2, 3, False, False, Ref3dPtg.ptg)]).stringify(self.wb), 'Sheet2!R3C4')
