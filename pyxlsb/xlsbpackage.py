import os
import shutil
from tempfile import TemporaryFile
from zipfile import ZipFile


class BasePackage(object):
    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def get_file(self, name):
        raise NotImplementedError

    def close(self):
        pass


class ZipPackage(BasePackage):
    def __init__(self, name):
        self._zf = ZipFile(name, 'r')

    def get_file(self, name):
        tf = TemporaryFile()
        try:
            with self._zf.open(name, 'r') as zf:
                shutil.copyfileobj(zf, tf)
            tf.seek(0, os.SEEK_SET)
            return tf
        except KeyError:
            tf.close()
            return None

    def close(self):
        self._zf.close()


class WorkbookPackage(BasePackage):
    def get_workbook_part(self):
        raise NotImplementedError

    def get_workbook_rels(self):
        raise NotImplementedError

    def get_sharedstrings_part(self):
        raise NotImplementedError

    def get_styles_part(self):
        raise NotImplementedError


class XlsbPackage(ZipPackage, WorkbookPackage):
    def get_workbook_part(self):
        return self.get_file('xl/workbook.bin')

    def get_workbook_rels(self):
        return self.get_file('xl/_rels/workbook.bin.rels')

    def get_sharedstrings_part(self):
        return self.get_file('xl/sharedStrings.bin')

    def get_styles_part(self):
        return self.get_file('xl/styles.bin')
