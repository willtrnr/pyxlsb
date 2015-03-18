import array
import os
import struct
import biff12
from cStringIO import StringIO
from handler import Handler, BasicHandler, StringTableHandler, StringInstanceHandler, SheetHandler, RowHandler, CellHandler

uint32_t = struct.Struct('I')
uint16_t = struct.Struct('H')
uint8_t = struct.Struct('B')
double_t = struct.Struct('d')

class RecordReader(object):
  def __init__(self, buf):
    self._fp = StringIO(buf)

  def skip(self, size):
    self._fp.seek(size, os.SEEK_CUR)

  def read(self, size):
    return self._fp.read(size)

  def read_int(self):
    buff = self._fp.read(4)
    if len(buff) < 4:
      return None
    buff = array.array('B', buff)
    buff.byteswap()
    return uint32_t.unpack(buff.tostring())[0]

  def read_short(self):
    buff = self._fp.read(2)
    if len(buff) < 2:
      return None
    buff = array.array('B', buff)
    buff.byteswap()
    return uint16_t.unpack(buff.tostring())[0]

  def read_byte(self):
    byte = self._fp.read(1)
    if byte == '':
      return None
    return uint8_t.unpack(byte)[0]

  def read_float(self):
    buff = self._fp.read(4)
    if len(buff) < 4:
      return None
    v = 0.0
    intval = uint32_t.unpack(buff)[0]
    if intval & 0x02 != 0:
      v = float(intval >> 2)
    else:
      buff = array.array('B', ('\x00' * 4) + uint32_t.pack(intval & 0xFFFFFFFC))
      v = double_t.unpack(buff.tostring())[0]
    if intval & 0x01 != 0:
      v /= 100
    return v

  def read_double(self):
    buff = self._fp.read(8)
    if len(buff) < 8:
      return None
    return double_t.unpack(buff)[0]

  def read_string(self):
    s = u''
    l = self.read_int()
    if l is None:
      return None
    for i in xrange(l):
      c = self.read_short()
      if c is None:
        return None
      s += unichr(c)
    return s


class BIFF12Reader(object):
  handlers = {
    # Workbook part handlers
    biff12.WORKBOOK:   BasicHandler('workbook'),
    biff12.SHEETS:     BasicHandler('sheets'),
    biff12.SHEETS_END: BasicHandler('/sheets'),
    biff12.SHEET:      SheetHandler(),

    # SharedStrings part handlers
    biff12.SST:     StringTableHandler(),
    biff12.SST_END: BasicHandler('/sst'),
    biff12.SI:      StringInstanceHandler(),

    # Worksheet part handlers
    biff12.WORKSHEET:     BasicHandler('worksheet'),
    biff12.WORKSHEET_END: BasicHandler('/worksheet'),
    biff12.SHEETDATA:     BasicHandler('sheetData'),
    biff12.SHEETDATA_END: BasicHandler('/sheetData'),
    # Sheet data cells handlers
    biff12.ROW:             RowHandler(),
    biff12.BLANK:           CellHandler(),
    biff12.NUM:             CellHandler(),
    biff12.BOOLERR:         CellHandler(),
    biff12.BOOL:            CellHandler(),
    biff12.FLOAT:           CellHandler(),
    biff12.STRING:          CellHandler(),
    biff12.FORMULA_STRING:  CellHandler(),
    biff12.FORMULA_FLOAT:   CellHandler(),
    biff12.FORMULA_BOOL:    CellHandler(),
    biff12.FORMULA_BOOLERR: CellHandler()
  }

  def __init__(self, fp):
    super(BIFF12Reader, self).__init__()
    self.debug = False
    self._fp = fp

  def __iter__(self):
    return self

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def read_id(self):
    v = 0
    for i in xrange(4):
      byte = self._fp.read(1)
      if byte == '':
        return None
      byte = uint8_t.unpack(byte)[0]
      v += byte << 8 * i
      if byte & 0x80 == 0:
        break
    return v

  def read_len(self):
    v = 0
    for i in xrange(4):
      byte = self._fp.read(1)
      if byte == '':
        return None
      byte = uint8_t.unpack(byte)[0]
      v += (byte & 0x7F) << (7 * i)
      if byte & 0x80 == 0:
        break
    return v

  def register_handler(self, recid, handler):
    self.handlers[recid] = handler

  def next(self):
    ret = None
    while ret is None:
      recid = self.read_id()
      reclen = self.read_len()
      if recid is None or reclen is None:
        raise StopIteration

      ret = (self.handlers.get(recid) or Handler()).read(RecordReader(self._fp.read(reclen)), recid, reclen)

      if self.debug:
        print '{:08X} {:04X} {:>8} {}'.format(self._fp.tell() - 8 - reclen, recid, reclen, ret)
    return (recid, ret)

  def close(self):
    self._fp.close()
