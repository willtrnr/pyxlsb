import unittest

from pyxlsb.cellreference import CellReference


class CellReferenceTestCase(unittest.TestCase):
    def test_col_to_index(self):
        self.assertEqual(CellReference.col_to_index("A"), 0)
        self.assertEqual(CellReference.col_to_index("B"), 1)
        self.assertEqual(CellReference.col_to_index("Z"), 25)
        self.assertEqual(CellReference.col_to_index("AA"), 26)
        self.assertEqual(CellReference.col_to_index("AB"), 27)
        self.assertEqual(CellReference.col_to_index("AZ"), 51)
        self.assertEqual(CellReference.col_to_index("ZZ"), 701)
        self.assertEqual(CellReference.col_to_index("XFD"), 16383)

    def test_index_to_col(self):
        self.assertEqual(CellReference.index_to_col(0), "A")
        self.assertEqual(CellReference.index_to_col(1), "B")
        self.assertEqual(CellReference.index_to_col(25), "Z")
        self.assertEqual(CellReference.index_to_col(26), "AA")
        self.assertEqual(CellReference.index_to_col(27), "AB")
        self.assertEqual(CellReference.index_to_col(51), "AZ")
        self.assertEqual(CellReference.index_to_col(701), "ZZ")
        self.assertEqual(CellReference.index_to_col(16383), "XFD")

    def test_str(self):
        self.assertEqual(str(CellReference(0, 0)), "$A$1")
        self.assertEqual(str(CellReference(0, 0, True, True)), "A1")

    def test_parse(self):
        ref = CellReference.parse("$A$1")
        self.assertEqual(ref.row, 0)
        self.assertEqual(ref.col, 0)
        self.assertFalse(ref.row_rel)
        self.assertFalse(ref.col_rel)

        ref = CellReference.parse("A1")
        self.assertEqual(ref.row, 0)
        self.assertEqual(ref.col, 0)
        self.assertTrue(ref.row_rel)
        self.assertTrue(ref.col_rel)
