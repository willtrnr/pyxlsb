import pyxlsb.records as records
from pyxlsb.recordreader import RecordReader


class Styles(object):
    def __init__(self, fp):
        self._reader = RecordReader(fp)
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self._colors = list()
        self._dxfs = list()
        self._table_styles = list()
        self._fills = list()
        self._fonts = list()
        self._borders = list()
        self._cell_xfs = list()
        self._cell_styles = list()
        self._cell_style_xfs = list()

        for rectype, reclen in self._reader:
            # TODO
            if rectype == records.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        return None

    def close(self):
        self._reader.close()
