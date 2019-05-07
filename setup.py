import os.path
from setuptools import setup
from subprocess import Popen, PIPE

# Get a handy base dir
project_dir = os.path.abspath(os.path.dirname(__file__))

# Use pandoc to convert the Markdown readme to RST for prtty printing on PyPI
try:
  proc = Popen(['pandoc', '--from', 'markdown', '--to', 'rst', '--output', '-', os.path.join(project_dir, 'README.md')], stdout=PIPE)
  README = proc.stdout.read().decode('utf-8')
  proc.stdout.close()
  proc.wait()
except:
  README = ''

setup(
  name='pyxlsb',
  version='1.0.5',

  description='Excel 2007-2010 Binary Workbook (xlsb) parser',
  long_description=README,

  author='William Turner',
  author_email='willtur.will@gmail.com',

  url='https://github.com/wwwiiilll/pyxlsb',

  license='LGPLv3+',

  classifiers=[
    'Development Status :: 5 - Production/Stable',

    'License :: OSI Approved :: GNU Lesser General Public License v3 or later (LGPLv3+)',

    'Programming Language :: Python :: 2',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.3',
    'Programming Language :: Python :: 3.4',
    'Programming Language :: Python :: 3.5',
    'Programming Language :: Python :: 3.6'
  ],

  packages=['pyxlsb']
)
