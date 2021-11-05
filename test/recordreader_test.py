import unittest

from pyxlsb.recordreader import RecordReader


class RecordReaderTestCase(unittest.TestCase):
    def test_read_type(self):
        reader = RecordReader(b"\xFD\x04\x00")
        self.assertEqual(reader._read_type(), 637)
        self.assertEqual(reader._read_type(), 0)
        self.assertEqual(reader._read_type(), None)

    def test_read_len(self):
        reader = RecordReader(b"\xC8\x01\x6C\x01")
        self.assertEqual(reader._read_len(), 200)
        self.assertEqual(reader._read_len(), 108)
        self.assertEqual(reader._read_len(), 1)
        self.assertEqual(reader._read_len(), None)
