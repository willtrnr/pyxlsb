from . import records
from .record_reader import RecordReader

class StringTable(object):
    def __init__(self, fp, _debug=False):
        super(StringTable, self).__init__()
        self._reader = RecordReader(fp=fp, _debug=_debug)
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
            if recid == records.SI:
                self._strings.append(rec.t)
            elif recid == records.SST_END:
                break

    def get_string(self, idx):
        return self._strings[idx]

    def close(self):
        self._reader.close()
