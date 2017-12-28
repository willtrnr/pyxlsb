import os.path
import unittest
from pyxlsb.workbook import Workbook
from pyxlsb.xlsbpackage import XlsbPackage

class WorkbookTestCase(unittest.TestCase):
    def setUp(self):
        self.workbook = Workbook(XlsbPackage(os.path.join('test_files', 'Test.xlsb')))

    def tearDown(self):
        self.workbook.close()

    def test_sheets(self):
        self.assertEqual(self.workbook.sheets, ['Test'])
