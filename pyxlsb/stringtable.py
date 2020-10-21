from . import biff12
from .reader import BIFF12Reader
from .writer import BIFF12Writer
from .handlers import StringInstanceHandler

class StringTable(object):
  def __init__(self, fp, prior_string_table=None, mode='r', debug=False):
    super(StringTable, self).__init__()
    self._mode = mode
    if mode == 'r':
      self._reader = BIFF12Reader(fp=fp, debug=debug)
    else:
      self._writer = BIFF12Writer(fp=fp, debug=debug)
    self._strings = list()
    if self._mode == 'r':
      self._parse()
    else:
      if prior_string_table is not None:
        self._count = prior_string_table._count
        self._unique_count = prior_string_table._unique_count
        self._strings = prior_string_table._strings.copy()
      else:
        self._count = self._unique_count = 0
        self._strings = []

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def __getitem__(self, key):
    return self._strings[key]

  def _parse(self):
    assert self._mode == 'r', 'Not opened for reading'
    for item in self._reader:
      if item[0] == biff12.SI:
        self._strings.append(item[1].t)
      elif item[0] == biff12.SST_END:
        break
      elif item[0] == biff12.SST:
        self._count = item[1].count
        self._unique_count = item[1].uniqueCount

  def get_string(self, idx):
    return self._strings[idx]

  def close(self):
    try:
      self._reader.close()
    except AttributeError:
      self._writer.close()

  def update(self, more_strings):
    known_strings = set(self._strings)
    for another_str in more_strings:
      if another_str in known_strings:
        self._count += 1
      else:
        self._strings.append(another_str)
        known_strings.add(another_str)
        self._unique_count += 1
        self._count += 1

  def write_table(self):
    writer = self._writer
    writer.write_id(biff12.SST)
    writer._writer.write_len(8)  # Always 2 integers.
    writer._writer.write_int(self._count)
    writer._writer.write_int(self._unique_count)
    for each_string in self._strings:
      writer.write_id(biff12.SI)
      StringInstanceHandler.write(writer, each_string)
    writer.write_id(biff12.SST_END)
    writer._writer.write_len(0)
