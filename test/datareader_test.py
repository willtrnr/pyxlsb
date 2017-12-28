import unittest
from pyxlsb.datareader import DataReader

class DataReaderTestCase(unittest.TestCase):
    def test_read_byte(self):
        reader = DataReader(b'\x01\x02\xff')
        self.assertEqual(reader.read_byte(), 0x01)
        self.assertEqual(reader.read_byte(), 0x02)
        self.assertEqual(reader.read_byte(), 0xFF)
        self.assertEqual(reader.read_byte(), None)

    def test_read_short(self):
        reader = DataReader(b'\x01\x00\x02\x00\x00\xff')
        self.assertEqual(reader.read_short(), 0x0001)
        self.assertEqual(reader.read_short(), 0x0002)
        self.assertEqual(reader.read_short(), 0xFF00)
        self.assertEqual(reader.read_short(), None)

    def test_read_int(self):
        reader = DataReader(b'\x01\x00\x00\x00\x02\x00\x00\x00\x00\xff\x00\xff')
        self.assertEqual(reader.read_int(), 0x00000001)
        self.assertEqual(reader.read_int(), 0x00000002)
        self.assertEqual(reader.read_int(), 0xFF00FF00)
        self.assertEqual(reader.read_int(), None)

    @unittest.skip('TODO')
    def test_read_float(self):
        pass

    @unittest.skip('TODO')
    def test_read_double(self):
        pass

    @unittest.skip('TODO')
    def test_read_rk(self):
        pass

    def test_read_string(self):
        reader = DataReader(b'\x04\x00\x00\x00T\x00E\x00S\x00T\x00')
        self.assertEqual(reader.read_string(), 'TEST')
        self.assertEqual(reader.read_string(), None)
