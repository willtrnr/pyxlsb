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
  return datetime(1900, 1, 1, 0, 0, 0) + timedelta(days=long(date), seconds=long((date % 1) * 24 * 60 * 60))
