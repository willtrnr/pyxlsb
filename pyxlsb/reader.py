import os
import struct
from io import BytesIO
from . import records

uint8_t = struct.Struct('<B')
uint16_t = struct.Struct('<H')
int32_t = struct.Struct('<i')
uint32_t = struct.Struct('<I')
float_t = struct.Struct('<f')
double_t = struct.Struct('<d')

class DataReader(object):
    def __init__(self, buf, enc='latin-1'):
        self._fp = buf if hasattr(buf, 'read') else BytesIO(buf)
        self._enc = enc

    def tell(self):
        return self._fp.tell()

    def seek(self, offset, whence=os.SEEK_SET):
        self._fp.seek(offset, whence)

    def skip(self, size):
        self._fp.seek(size, os.SEEK_CUR)

    def read(self, size):
        return self._fp.read(size)

    def read_byte(self):
        byte = self._fp.read(1)
        if byte == b'':
            return None
        return uint8_t.unpack(byte)[0]

    def read_short(self):
        buff = self._fp.read(2)
        if len(buff) < 2:
            return None
        return uint16_t.unpack(buff)[0]

    def read_int(self):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        return uint32_t.unpack(buff)[0]

    def read_float(self):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        return float_t.unpack(buff)[0]

    def read_double(self):
        buff = self._fp.read(8)
        if len(buff) < 8:
            return None
        return double_t.unpack(buff)[0]

    def read_string(self):
        l = self.read_int()
        if l is None:
            return None
        buff = self.read(l * 2)
        if len(buff) < l * 2:
            return None
        return buff.decode(self._enc).replace('\x00', '')

    def read_rk(self):
        buff = self._fp.read(4)
        if len(buff) < 4:
            return None
        v = 0.0
        intval = int32_t.unpack(buff)[0]
        if intval & 0x02 != 0:
            v = float(intval >> 2)
        else:
            v = double_t.unpack(b'\x00\x00\x00\x00' + uint32_t.pack(intval & 0xFFFFFFFC))[0]
        if intval & 0x01 != 0:
            v /= 100
        return v
