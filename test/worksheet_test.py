import os.path
import unittest
from unittest.mock import Mock

from pyxlsb.worksheet import Worksheet


class WorksheetTestCase(unittest.TestCase):
    def setUp(self):
        wb = Mock()
        wb.get_shared_string = Mock(return_value="SS")
        self.sheet = Worksheet(
            wb, "Sheet1", open(os.path.join("test_files", "parts", "sheet1.bin"), "rb")
        )

    def tearDown(self):
        self.sheet.close()

    def test_dimension(self):
        self.assertNotEqual(self.sheet.dimension, None)
        self.assertEqual(self.sheet.dimension.r, 0)
        self.assertEqual(self.sheet.dimension.c, 0)
        self.assertEqual(self.sheet.dimension.h, 5)
        self.assertEqual(self.sheet.dimension.w, 8)

    def test_rows(self):
        rows = list(self.sheet.rows())

        # We should have as much rows as dim.h
        self.assertEqual(len(rows), 5)
        for rn, r in enumerate(rows):
            # Each row should have as much cells as dim.w
            for cn, c in enumerate(r):
                # Each cell should have the proper row and cell number
                self.assertEqual(c.r, rn)
                self.assertEqual(c.c, cn)
