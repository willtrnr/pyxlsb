import os.path
import unittest
from pyxlsb.stringtable import StringTable

class StringTableTestCase(unittest.TestCase):
    def setUp(self):
        self.stringtable = StringTable(open(os.path.join('test_files', 'parts', 'sharedStrings.bin'), 'rb'))

    def tearDown(self):
        self.stringtable.close()

    def test_get_string(self):
        self.assertEqual(self.stringtable.get_string(0), 'A')
        self.assertEqual(self.stringtable.get_string(1), 'B')
