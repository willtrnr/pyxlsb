from . import biff12
from .reader import uint8_t, uint16_t, int32_t, uint32_t, double_t, BIFF12Reader
import os
try:
    import numpy as np
except ImportError:
    np = None


class RecordWriter(object):
    def __init__(self, fp, enc='utf-16'):
        self._fp = fp
        self._enc = enc

    def tell(self):
        return self._fp.tell()

    def seek(self, offset, whence=os.SEEK_SET):
        self._fp.seek(offset, whence)

    def skip(self, size):
        self._fp.seek(size, os.SEEK_CUR)

    def write(self, data):
        return self._fp.write(data)

    def write_int(self, val, do_write_len=False):  # TODO: remove do_write_len option altogether?
        payload = uint32_t.pack(val)
        if do_write_len:
            self.write_len(len(payload))
        self._fp.write(payload)

    def write_short(self, val, do_write_len=False):
        payload = uint16_t.pack(val)
        if do_write_len:
            self.write_len(len(payload))
        self._fp.write(payload)

    def write_byte(self, val, do_write_len=False):
        payload = uint8_t.pack(val)
        if do_write_len:
            self.write_len(len(payload))
        self._fp.write(payload)

    def write_float(self, val, do_write_len=False):
        payload = struct.Struct('<f').pack(val)

        # TODO: finish adjustment per read_float steps below:
        ### v = 0.0
        ### intval = int32_t.unpack(buff)[0]
        ### if intval & 0x02 != 0:
        ###     v = float(intval >> 2)
        ### else:
        ###     v = double_t.unpack(b'\x00\x00\x00\x00' + uint32_t.pack(intval & 0xFFFFFFFC))[0]
        ### if intval & 0x01 != 0:
        ###     v /= 100
        ### return v

        if do_write_len:
            self.write_len(len(payload))
        self._fp.write(payload)

    def write_double(self, val, do_write_len=False):
        payload = double_t.pack(val)
        if do_write_len:
            self.write_len(len(payload))
        self._fp.write(payload)

    def write_string(self, str_data):
        try:
            data = str_data.encode(self._enc, errors='replace')
            size = len(data) // 2
            self.write_int(size, do_write_len=False)  # TODO: Matches RecordReader.read_string but not RecordReader.read_len?
            self._fp.write(data)
        except AttributeError:
            # Reference to shared string was passed in instead.
            self.write_int(val=str_data)

    def write_obj_str(self, obj):
        try:
            assert not isinstance(obj, int)
            data = str(obj).encode(self._enc, errors='replace')
            size = len(data) // 2
            self.write_int(size, do_write_len=False)  # TODO: Matches RecordReader.read_string but not RecordReader.read_len?
            self._fp.write(data)
        except AssertionError:
            # Reference to shared string was passed in instead.
            self.write_int(val=obj)

    def write_len(self, val):
        # TODO: Does not match what RecordReader.read_string does? (Does match RecordReader.read_len)
        for i in range(4):
            nearlybyte = (val >> 7 * i) & 0x7F
            payload = uint8_t.pack(nearlybyte)
            self._fp.write(payload)
            if nearlybyte & 0x80 == 0:
                break


class SemiflexibleHandlers(dict):

    def __init__(self, *args, **kwargs):
        initial_handlers = {
            int: (biff12.NUM, RecordWriter.write_double),  # TODO: can/should support long?
            float: (biff12.NUM, RecordWriter.write_double),
            bool: (biff12.BOOL, RecordWriter.write_byte),
            str: (biff12.STRING, RecordWriter.write_string),
        }
        dict.__init__(self, initial_handlers, *args, **kwargs)
        if np is not None:
            self.update(
                {
                    np.dtype('int64'): (biff12.FLOAT, RecordWriter.write_double),   # TODO: can support long in NUM?
                    np.dtype('int32'): (biff12.NUM, RecordWriter.write_int),
                    np.dtype('int16'): (biff12.NUM, RecordWriter.write_short),      # TODO: write_int? but then is write_short ever used?
                    np.dtype('int8'): (biff12.NUM, RecordWriter.write_byte),        # TODO: write_int?
                    np.dtype('float128'): (biff12.FLOAT, RecordWriter.write_double),
                    np.dtype('float64'): (biff12.FLOAT, RecordWriter.write_double),
                    np.dtype('float32'): (biff12.FLOAT, RecordWriter.write_float),  # TODO: write_double?
                    np.dtype('float16'): (biff12.FLOAT, RecordWriter.write_float),  # TODO: write_double?
                    np.dtype('bool8'): (biff12.BOOL, RecordWriter.write_byte),
                    np.dtype('object'): (biff12.STRING, RecordWriter.write_obj_str),
                }
            )

    def __missing__(self, key):
        if isinstance(key, int):
            value = self[int]
            self[key] = value
        elif isinstance(key, float):
            value = self[float]
            self[key] = value
        elif isinstance(key, bool):
            value = self[bool]
            self[key] = value
        elif isinstance(key, str):
            value = self[str]
            self[key] = value
        else:
            value = (biff12.STRING, RecordWriter.write_obj_str)
            self[key] = value
        return value


class BIFF12Writer(BIFF12Reader):

    handlers = SemiflexibleHandlers()

    def __init__(self, fp, debug=False):
        self._debug = debug
        self._writer = RecordWriter(fp=fp)  # TODO: Move to Worksheet class most likely
        self._fp = self._writer._fp

    def read_id(self):
        raise NotImplemented  # Inherited but not desired

    def read_len(self):
        raise NotImplemented  # Inherited but not desired

    def next(self):
        raise StopIteration   # Inherited but not desired; TODO: new shared base class
    #    ret = None
    #    while ret is None:
    #      if self._debug:
    #        pos = self._fp.tell()
    #      recid = self.read_id()
    #      reclen = self.read_len()
    #      if recid is None or reclen is None:
    #        raise StopIteration
    #      recdata = self._fp.read(reclen)
    #      ret = (self.handlers.get(recid) or Handler()).read(RecordReader(recdata), recid, reclen)
    #      if self._debug:
    #        print('{:08X}  {:04X}  {:<6} {} {}'.format(pos, recid, reclen, ' '.join('{:02X}'.format(b) for b in recdata), ret))
    #    return (recid, ret)

    def write_id(self, val):
        # TODO: optional assert val is in allowed values from biff12?
        _write = self._writer.write
        for i in range(4):
            byte = (val >> 8 * i) & 0xFF
            payload = uint8_t.pack(byte)
            _write(payload)
            if byte & 0x80 == 0:
                break

    def write_string(self, str_data):
        self._writer.write_string(str_data)

    def write_bytes(self, bytes_data):
        assert isinstance(bytes_data, bytes)
        self._writer.write_len(len(bytes_data))
        if len(bytes_data):
            self._writer.write(bytes_data)


test_code = '''
import pyxlsb
wb = pyxlsb.open_workbook('s.xlsb', 'w')
ws = wb.create_sheet('blah')
import pandas as pd
df = pd.DataFrame( 
    [ 
        ['blue', 4, -273.154], 
        ['azul', 5, -195.79], 
        ['blau', 6, -78.5], 
    ], 
    columns=['colors', 'points', 'temp'] 
)                                                                       
ws.write_table(df)
ws.close()
wb.close()
wb = pyxlsb.open_workbook('s.xlsb', debug=True)
ws = wb.get_sheet(1)
for row in ws.rows():
    print(row)
'''
