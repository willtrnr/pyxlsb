from biff12 import *
from handlers import Handler
from reader import BIFF12Reader
from workbook import Workbook
from worksheet import Worksheet

def open_workbook(name):
  from zipfile import ZipFile
  zf = ZipFile(name, 'r')
  return Workbook(fp=zf)

def convert_date(date):
  if not isinstance(date, int) and not isinstance(date, long) and not isinstance(date, float):
    return None

  from datetime import datetime, timedelta
  if long(date) == 0:
    return datetime(1900, 1, 1, 0, 0, 0) + timedelta(seconds=date * 24 * 60 * 60)
  elif date >= 61.0:
    # According to Lotus 1-2-3, Feb 29th 1900 is a real thing, therefore we have to remove one day after that date
    return datetime(1899, 12, 31, 0, 0, 0) + timedelta(days=long(date) - 1, seconds=long((date % 1) * 24 * 60 * 60))
  else:
    # Feb 29th 1900 will show up as Mar 1st 1900 because Python won't handle that date
    return datetime(1899, 12, 31, 0, 0, 0) + timedelta(days=long(date), seconds=long((date % 1) * 24 * 60 * 60))
