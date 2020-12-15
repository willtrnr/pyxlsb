import os.path
import unittest
from datetime import datetime, time
from pyxlsb.workbook import Workbook
from pyxlsb.xlsbpackage import XlsbPackage


class WorkbookTestCase(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook(XlsbPackage(os.path.join('test_files', 'test.xlsb')))
        self.wb1904 = Workbook(XlsbPackage(os.path.join('test_files', 'test_1904.xlsb')))
        self.wb_chart = Workbook(XlsbPackage(os.path.join('test_files', 'test_chart_sheet.xlsb')))

    def tearDown(self):
        self.wb.close()
        self.wb1904.close()
        self.wb_chart.close()

    def test_props(self):
        self.assertNotEqual(self.wb.props, None)
        self.assertEqual(self.wb.props.date1904, False)

    def test_props_1904(self):
        self.assertNotEqual(self.wb1904.props, None)
        self.assertEqual(self.wb1904.props.date1904, True)

    def test_sheets(self):
        self.assertEqual(self.wb.sheets, ['Test'])
        self.assertEqual(self.wb_chart.sheets, ['Sheet1', 'Chart1', 'Sheet3'])

    def test_convert_date(self):
        # Both 0 and 1 are the same date, only 0 is used for time only values
        self.assertEqual(self.wb.convert_date(0), datetime(1900, 1, 1))
        self.assertEqual(self.wb.convert_date(1), datetime(1900, 1, 1))

        # Check for the special 1900-02-29 case from Lotus 1-2-3
        self.assertEqual(self.wb.convert_date(59), datetime(1900, 2, 28))
        self.assertEqual(self.wb.convert_date(60), datetime(1900, 3, 1))
        self.assertEqual(self.wb.convert_date(61), datetime(1900, 3, 1))
        self.assertEqual(self.wb.convert_date(62), datetime(1900, 3, 2))

        # Test date with time component
        self.assertEqual(self.wb.convert_date(1.25), datetime(1900, 1, 1, 6, 0, 0))

    def test_convert_date_1904(self):
        # Test that 0 and 1 are the same date, 0 is used for time only value
        self.assertEqual(self.wb1904.convert_date(0), datetime(1904, 1, 1))
        self.assertEqual(self.wb1904.convert_date(1), datetime(1904, 1, 1))

        # Make sure we don't apply the 1900-02-29 special case
        self.assertEqual(self.wb1904.convert_date(60), datetime(1904, 2, 29))
        self.assertEqual(self.wb1904.convert_date(61), datetime(1904, 3, 1))
        self.assertEqual(self.wb1904.convert_date(62), datetime(1904, 3, 2))

        # Test date with time component
        self.assertEqual(self.wb1904.convert_date(1.25), datetime(1904, 1, 1, 6, 0, 0))

    def test_convert_time(self):
        self.assertEqual(self.wb.convert_time(0), time(0, 0, 0))
        self.assertEqual(self.wb.convert_time(0.25), time(6, 0, 0))
