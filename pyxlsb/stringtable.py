from . import biff12
from .reader import BIFF12Reader

class StringTable(object):
    def __init__(self, fp):
        super(StringTable, self).__init__()
        self._reader = BIFF12Reader(fp=fp)
        self._strings = list()
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def __getitem__(self, key):
        return self._strings[key]

    def _parse(self):
        for recid, rec in self._reader:
            if recid == biff12.SI:
                self._strings.append(rec.t)
            elif recid == biff12.SST_END:
                break

    def get_string(self, idx):
        return self._strings[idx]

    def close(self):
        self._reader.close()
