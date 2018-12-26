import os
import struct
import sys
from io import BytesIO

if sys.version_info > (3,):
    xrange = range

_uint8_t = struct.Struct('<B')
_uint16_t = struct.Struct('<H')
_int32_t = struct.Struct('<i')
_uint32_t = struct.Struct('<I')
_float_t = struct.Struct('<f')
_double_t = struct.Struct('<d')


class DataReader(object):
    def __init__(self, fp, enc='utf-16'):
        if isinstance(fp, DataReader):
            self._fp = fp._fp
        elif hasattr(fp, 'read') and hasattr(fp, 'seek'):
            self._fp = fp
        else:
            self._fp = BytesIO(fp)
        self._enc = enc

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def tell(self):
        return self._fp.tell()

    def seek(self, offset, whence=os.SEEK_SET):
        self._fp.seek(offset, whence)

    def skip(self, size):
        self._fp.seek(size, os.SEEK_CUR)

    def read(self, size):
        return self._fp.read(size)

    def read_bool(self):
        buf = self.read(1)
        if not buf:
            return None
        return buf == b'\x01'

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

        buf = self.read(size * 2)
        if len(buf) != size * 2:
            return None

        return buf.decode(enc or self._enc)

    def close(self):
        self._fp.close()
