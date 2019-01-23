import os
import glob
import sys
from setuptools import setup, find_packages

setupargs = {}

setup(name='exceltools',
      version='0.1.6',
      packages=find_packages('src'),
      package_dir={'': 'src'},

      # dependencies:
      install_requires = ['xlrd', 'xlsxwriter'],

      # PyPI metadata
      author='Danian Hu',
      author_email='hudanian@gmail.com',
      description='report in excel.',
      **setupargs
     )
