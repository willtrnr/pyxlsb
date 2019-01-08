import os
import struct
from io import BytesIO


_uint8_t = struct.Struct('<B')
_uint16_t = struct.Struct('<H')
_int32_t = struct.Struct('<i')
_uint32_t = struct.Struct('<I')
_float_t = struct.Struct('<f')
_double_t = struct.Struct('<d')


class DataReader(object):
    def __init__(self, buf, enc=None):
        self._buf = BytesIO(buf)
        self._enc = enc if enc is not None else 'utf-16'

    def skip(self, size):
        self._buf.seek(size, os.SEEK_CUR)

    def read(self, size):
        return self._buf.read(size)

    def read_bool(self):
        buf = self.read(1)
        if not buf:
            return None
        return buf != b'\x00'

    def read_byte(self):
        buf = self.read(1)
        if not buf:
            return None
        return _uint8_t.unpack(buf)[0]

    def read_short(self):
        buf = self.read(2)
        if len(buf) < 2:
            return None
        return _uint16_t.unpack(buf)[0]

    def read_int(self):
        buf = self.read(4)
        if len(buf) < 4:
            return None
        return _uint32_t.unpack(buf)[0]

    def read_float(self):
        buf = self.read(4)
        if len(buf) < 4:
            return None
        return _float_t.unpack(buf)[0]

    def read_double(self):
        buf = self.read(8)
        if len(buf) < 8:
            return None
        return _double_t.unpack(buf)[0]

    def read_rk(self):
        buf = self.read(4)
        if len(buf) < 4:
            return None

        value = 0.0
        intval = _int32_t.unpack(buf)[0]
        if intval & 0x02 == 0x02:
            value = float(intval >> 2)
        else:
            value = _double_t.unpack(b'\x00\x00\x00\x00' + _uint32_t.pack(intval & 0xFFFFFFFC))[0]

        if intval & 0x01 == 0x01:
            value /= 100

        return value

    def read_string(self, size=None, enc=None):
        if size is None:
            size = self.read_int()
            if size is None:
                return None

        size *= 2
        buf = self.read(size)
        if len(buf) != size:
            return None

        return buf.decode(enc or self._enc)
