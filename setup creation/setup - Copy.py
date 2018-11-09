from cx_Freeze import setup, Executable
import time
import sys
import win32com.client
import autoit
from selenium import webdriver
import csv
import re
import tkinter
from tkinter import *
import numpy
import selenium
#import pytz
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait as wait
#from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
#from selenium.webdriver.chrome.options import Options
#from pyvirtualdisplay import Display
#from xvfbwrapper import Xvfb
import xlwt
import csv

import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

#os.environ['TCL_LIBRARY'] = r'C:\Users\balaji.ma\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
#os.environ['TK_LIBRARY'] = r'C:\Users\balaji.ma\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

packages=['time','win32com.client','autoit','tkinter','csv','re','selenium','xlwt','numpy']
executables = [
    Executable('CRQStatus.py', base=base)
]

setup(name='CRQ Status',
      version='0.1',
      description='Sample CRQ status retrieval',
      options={'build_exe':{'packages' :packages, 'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
         ]

        }
               },
      executables=executables
      )
