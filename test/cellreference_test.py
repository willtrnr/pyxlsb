import unittest
from pyxlsb.cellreference import CellReference

class CellReferenceTestCase(unittest.TestCase):
    def test_col_to_index(self):
        self.assertEqual(CellReference.col_to_index('A'), 0)
        self.assertEqual(CellReference.col_to_index('B'), 1)
        self.assertEqual(CellReference.col_to_index('Z'), 25)
        self.assertEqual(CellReference.col_to_index('AA'), 26)
        # FIXME: Pretty sure this is not right
        self.assertEqual(CellReference.col_to_index('ZZ'), 701)
