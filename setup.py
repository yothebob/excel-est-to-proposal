from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

setup(
    options = {'py2exe': {'bundle_files': 1, 'compressed': True, 'includes': ['lxml.etree','lxml._elementpath', 'gzip','docx','openpyxl','datetime','re','os','tkinter','sys',]}},
    windows = [{'script': "proposal_writer.py"}],
    zipfile = None,
)
