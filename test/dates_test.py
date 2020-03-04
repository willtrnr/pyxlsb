import os.path
import unittest
from pyxlsb import open_workbook


class DatesTestCase(unittest.TestCase):
    def setUp(self):
        self.wb = open_workbook(os.path.join('test_files', 'dates.xlsb'))

    def tearDown(self):
        self.wb.close()

    def test_dates(self):
        with self.wb.get_sheet_by_index(0) as s:
            cells = [c for r in s for c in r]
        self.assertEqual(len(cells), 3)
        self.assertTrue(cells[0].is_date_formatted)
        self.assertTrue(cells[1].is_date_formatted)
        self.assertTrue(cells[2].is_date_formatted)
