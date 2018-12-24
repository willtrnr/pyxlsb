# -*- coding: utf-8 -*-
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

    def test_read_bool(self):
        reader = DataReader(b'\x00\x01\x02')
        self.assertEqual(reader.read_bool(), False)
        self.assertEqual(reader.read_bool(), True)
        # Not sure about this case
        self.assertEqual(reader.read_bool(), False)
        self.assertEqual(reader.read_bool(), None)

    # TODO: Add negative cases
    def test_read_rk(self):
        # RK type 0 IEEE
        reader = DataReader(b'\x00\x00\xF0\x3F')
        self.assertEqual(reader.read_rk(), 1)
        self.assertEqual(reader.read_rk(), None)

        # RK type 1 IEEE x 100
        reader = DataReader(b'\x01\xC0\x5E\x40')
        self.assertEqual(reader.read_rk(), 1.23)
        self.assertEqual(reader.read_rk(), None)

        # RK type 2 integer
        reader = DataReader(b'\x3A\x85\xF1\x02')
        self.assertEqual(reader.read_rk(), 12345678)
        self.assertEqual(reader.read_rk(), None)

        # RK type 3 integer x 100
        reader = DataReader(b'\x3B\x85\xF1\x02')
        self.assertEqual(reader.read_rk(), 123456.78)
        self.assertEqual(reader.read_rk(), None)

    def test_read_string(self):
        reader = DataReader(b'\x04\x00\x00\x00T\x00E\x00S\x00T\x00')
        self.assertEqual(reader.read_string(), 'TEST')
        self.assertEqual(reader.read_string(), None)

    def test_read_string_u(self):
        reader = DataReader(b'\x05\x00\x00\x00\x53\x30\x93\x30\x6B\x30\x61\x30\x6F\x30')
        self.assertEqual(reader.read_string(), u'こんにちは')
        self.assertEqual(reader.read_string(), None)
