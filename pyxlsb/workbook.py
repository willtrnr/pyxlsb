import os
import sys
from . import biff12
from .reader import BIFF12Reader
from .writer import BIFF12Writer
from .handlers import SheetHandler
from .stringtable import StringTable
from .worksheet import Worksheet, supported_modes
from tempfile import TemporaryFile

if sys.version_info > (3,):
  basestring = (str, bytes)

class Workbook(object):
  def __init__(self, fp, debug=False):
    super(Workbook, self).__init__()
    self._zf = fp
    self._mode = self._zf.mode
    assert self._mode in supported_modes
    self._debug = debug
    self.sheets = list()
    self.stringtable = None
    if self._mode == 'r':
      self._parse()
    else:
      try:
        with self._zf.open('xl/sharedStrings.bin', 'r') as shared_strings_zf:
          self.stringtable = StringTable(fp=shared_strings_zf, debug=self._debug)
      except KeyError:
        # xl/sharedStrings.bin file does not exist yet
        self.stringtable = StringTable(fp=None, mode='w', debug=self._debug)

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def _parse(self):
    assert self._mode == 'r', 'Not opened for reading'
    with TemporaryFile() as temp:
      with self._zf.open('xl/workbook.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      if self._debug:
        print('Parsing xl/workbook.bin')
      reader = BIFF12Reader(fp=temp, debug=self._debug)
      for item in reader:
        if item[0] == biff12.SHEET:
          self.sheets.append(item[1].name)
        elif item[0] == biff12.SHEETS_END:
          break

    try:
      temp = TemporaryFile()
      with self._zf.open('xl/sharedStrings.bin', 'r') as zf:
        temp.write(zf.read())
        temp.seek(0, os.SEEK_SET)
      if self._debug:
        print('Parsing xl/sharedStrings.bin')
      self.stringtable = StringTable(fp=temp)
    except KeyError:
      pass

  def get_sheet(self, idx, rels=False):
    assert self._mode == 'r', 'Not opened for reading'
    if isinstance(idx, basestring):
      idx = [s.lower() for s in self.sheets].index(idx.lower()) + 1
    if idx < 1 or idx > len(self.sheets):  # TODO: this check seems unnecessary given behavior of list.index
      raise IndexError('sheet index out of range')

    temp = TemporaryFile()
    with self._zf.open('xl/worksheets/sheet{}.bin'.format(idx), 'r') as zf:
      temp.write(zf.read())
      temp.seek(0, os.SEEK_SET)

    if rels:
      rels_temp = TemporaryFile()
      with self._zf.open('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx), 'r') as zf:
        rels_temp.write(zf.read())
        rels_temp.seek(0, os.SEEK_SET)
    else:
      rels_temp = None

    return Worksheet(fp=temp, rels_fp=rels_temp, stringtable=self.stringtable, debug=self._debug)

  def create_sheet(self, sheet_name):
    assert self._mode == 'w', 'Not opened for writing'
    assert sheet_name not in self.sheets, 'Duplicate sheet name'

    self.sheets.append(sheet_name)

    sheet_idx = len(self.sheets)
    sheet_zf = self._zf.open('xl/worksheets/sheet{}.bin'.format(sheet_idx), 'w')
    try:
      sheet = Worksheet(fp=sheet_zf, rels_fp=None, stringtable=self.stringtable, debug=self._debug, mode='w')
    except BaseException:
      sheet_zf.close()
      raise

    return sheet

  def _write_workbook_bin(self):
    _empty = b''
    with self._zf.open('xl/workbook.bin', 'w') as workbook_zf:
      writer = BIFF12Writer(fp=workbook_zf, debug=self._debug)
      writer.write_id(biff12.WORKBOOK)
      writer.write_bytes(_empty)
      writer.write_id(biff12.FILEVERSION)
      writer.write_bytes(b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x02\x00\x00\x00x\x00l\x00\x01\x00\x00\x005\x00\x01\x00\x00\x005\x00\x04\x00\x00\x009\x003\x000\x003\x00')
      writer.write_id(biff12.WORKBOOKPR)
      writer.write_bytes(b' \x00\x01\x00B\xe5\x01\x00\x00\x00\x00\x00')
      writer.write_id(biff12.BOOKVIEWS)
      writer.write_bytes(_empty)
      writer.write_id(biff12.WORKBOOKVIEW)
      writer.write_bytes(b'\xe0\x01\x00\x00x\x00\x00\x00\x13\x92\x00\x00#F\x00\x00X\x02\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00x')
      writer.write_id(biff12.BOOKVIEWS_END)
      writer.write_bytes(_empty)
      writer.write_id(biff12.SHEETS)
      writer.write_bytes(_empty)
      for sheetid, sheet_name in enumerate(self.sheets, start=1):
        writer.write_id(biff12.SHEET)
        relid = 'rId{}'.format(sheetid)  # TODO: properly connect up to rel stuff
        payload = SheetHandler.compose(writer._writer.__class__, sheetid, relid, sheet_name)
        writer.write_bytes(payload)
      writer.write_id(biff12.SHEETS_END)
      writer.write_bytes(_empty)
      # Ending nonsense
      writer._writer.write(b'\x9d\x01\x1a\xd5\x38\x02\x00\x01\x00\x00\x00\x64\x00\x00\x00\xfc\xa9\xf1\xd2\x4d\x62\x50\x3f\x02\x00\x00\x00\x6a\x00\x9b\x01\x01\x00\x84\x01\x00')

  def _write_shared_strings(self):
    with self._zf.open('xl/sharedStrings.bin', 'w') as shared_strings_zf:
      self.stringtable = StringTable(fp=shared_strings_zf, prior_string_table=self.stringtable, mode='w', debug=self._debug)
      self.stringtable.write_table()

  def _write_Content_Types_xml(self):
    with self._zf.open(r'[Content_Types].xml', 'w') as xml_zf:
      xml_zf.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>')
      for i in range(1, len(self.sheets)+1):
        xml_zf.write('<Override PartName="/xl/worksheets/sheet{}.bin" ContentType="application/vnd.ms-excel.worksheet"/>'.format(i).encode('utf8'))
      xml_zf.write(b'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/><Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')

  def _write_workbook_bin_rels(self):
    count_plus_1 = len(self.sheets)+1
    with self._zf.open('xl/_rels/workbook.bin.rels', 'w') as wbrels_zf:
      wbrels_zf.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/><Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'.format(count_plus_1 + 1, count_plus_1).encode('utf8'))
      for i in range(1, count_plus_1):
        wbrels_zf.write('<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.bin"/>'.format(i, i).encode('utf8'))
      wbrels_zf.write('<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/></Relationships>'.format(count_plus_1 + 2).encode('utf8'))

  def _write_rels_dot_rels(self):
    with self._zf.open('_rels/.rels', 'w') as relsrels_zf:
      relsrels_zf.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>')

  def close(self):
    if self._mode == 'w':
      self._write_workbook_bin()
      self._write_shared_strings()
      self._write_Content_Types_xml()
      self._write_workbook_bin_rels()
      self._write_rels_dot_rels()
    self._zf.close()
