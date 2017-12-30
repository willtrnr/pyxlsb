import os
import shutil
from tempfile import TemporaryFile
from zipfile import ZipFile

class ZipPackage(object):
    def __init__(self, name):
        self._zf = ZipFile(name, 'r')

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def get_file(self, name):
        tf = TemporaryFile()
        try:
            with self._zf.open(name, 'r') as zf:
                shutil.copyfileobj(zf, tf)
            tf.seek(0, os.SEEK_SET)
            return tf
        except KeyError:
            tf.close()
            raise

    def close(self):
        self._zf.close()

class XlsbPackage(ZipPackage):
    def get_workbook_part(self):
        return self.get_file('xl/workbook.bin')

    def get_sharedstrings_part(self):
        return self.get_file('xl/sharedStrings.bin')

    def get_styles_part(self):
        return self.get_file('xl/styles.bin')

    def get_worksheet_part(self, idx):
        return self.get_file('xl/worksheets/sheet{}.bin'.format(idx))

    def get_worksheet_rels(self, idx):
        try:
            return self.get_file('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx))
        except KeyError:
            return None
