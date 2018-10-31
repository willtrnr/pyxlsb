import os
import struct
from io import BytesIO

_int8_t = struct.Struct('<b')
_uint8_t = struct.Struct('<B')
_int16_t = struct.Struct('<h')
_uint16_t = struct.Struct('<H')
_int32_t = struct.Struct('<i')
_uint32_t = struct.Struct('<I')
_float_t = struct.Struct('<f')
_double_t = struct.Struct('<d')


class DataReader(object):
    def __init__(self, fp, enc='utf-16'):
        self._fp = fp if hasattr(fp, 'read') else BytesIO(fp)
        self._enc = enc

    def tell(self):
        return self._fp.tell()

    def seek(self, offset, whence=os.SEEK_SET):
        self._fp.seek(offset, whence)

    def skip(self, size):
        self._fp.seek(size, os.SEEK_CUR)

    def read(self, size):
        return self._fp.read(size)

    def read_byte(self, signed=False):
        byte = self._fp.read(1)
        if not byte:
            return None
        return (_int8_t if signed else _uint8_t).unpack(byte)[0]

    def read_short(self, signed=False):
        buff = self._fp.read(2)
        if len(buff) < 2:
            return None
        return (_int16_t if signed else _uint16_t).unpack(buff)[0]

    def read_int(self, signed=False):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        return (_int32_t if signed else _uint32_t).unpack(buff)[0]

    def read_float(self):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        return _float_t.unpack(buff)[0]

    def read_double(self):
        buff = self._fp.read(8)
        if len(buff) < 8:
            return None
        return _double_t.unpack(buff)[0]

    def read_bool(self):
        value = self._fp.read(1)
        if not value:
            return None
        return value == b'\x01'

    def read_rk(self):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        value = 0.0
        intval = _int32_t.unpack(buff)[0]
        if intval & 0x02 == 0x02:
            value = float(intval >> 2)
        else:
            value = _double_t.unpack(b'\x00\x00\x00\x00' + _uint32_t.pack(intval & 0xFFFFFFFC))[0]
        if intval & 0x01 == 0x01:
            value /= 100
        return value

    def read_string(self, size=None, enc=None):
        size = size or self.read_int()
        if size is None:
            return None
        buff = self.read(size * 2)
        if len(buff) != size * 2:
            return None
        return buff.decode(enc or self._enc)
