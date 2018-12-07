from cx_Freeze import setup, Executable

import time
import sys
import win32com.client
import csv
import re
import tkinter
from tkinter import *
#import numpy
import selenium
import win32api

import sqlite3
#from sqlite3 import *
import shutil
import win32api
import datetime

import os



import pandas_datareader
import pandas
#import os
#import win32com.client
import plotly
import numpy
#get_ipython().magic(u'matplotlib inline')
import seaborn
import matplotlib
import cufflinks
    



PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

#os.environ['TCL_LIBRARY'] = r'C:\Users\balaji.ma\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
#os.environ['TK_LIBRARY'] = r'C:\Users\balaji.ma\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

packages=[
          #'final_flow',
          #'creating_db',
          #'copy_paste',
          #'Backlog_report_generator_loop',
          #'ui_file',
          'sqlite3',
          'shutil',
          'win32api',
          'datetime',          
          'time','win32com.client','tkinter','csv','re','selenium','win32api',
          'plotly','matplotlib','cufflinks',
        'numpy','pandas_datareader','pandas']
executables = [
    Executable('backlog_report.py', base=base)
]

setup(name='Backlog report automation',
      version='0.1',
      description='Backlog report automation',
      options={'build_exe':{'packages' :packages, 'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
            os.path.join(sys.base_prefix, 'DLLs', 'sqlite3.dll'),
            'C:\\Users\\balaji.ma\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pypiwin32_system32\\pythoncom36.dll',
            'C:\\Users\\balaji.ma\\AppData\\Local\\Programs\\Python\\Python36-32\\Lib\\site-packages\\pypiwin32_system32\\pywintypes36.dll',
            'C:\\Users\\balaji.ma\\AppData\\Local\Programs\\Python\\Python36-32\\Lib\\site-packages\\win32\\pythoncom36.dll',
            'C:\\Users\\balaji.ma\\AppData\\Local\Programs\\Python\\Python36-32\\Lib\\site-packages\\win32\\pywintypes36.dll',
         ]

        }
               },
      executables=executables
      )
