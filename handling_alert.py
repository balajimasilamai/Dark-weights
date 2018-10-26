from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
#import os
import win32com.client

import datetime

def download_file():
    #Selenium part to download the files
    #with VPN
    #options = Options()
    #options.add_argument("--disable-notifications")
    driver=webdriver.Chrome()#chrome_options=options
    driver.maximize_window()
    window_before = driver.window_handles[0]
    try:
        driver.get('https://itsm.windstream.com/')
        time.sleep(20)
        #WebDriverWait(driver,60)
        #pythoncom.CoInitialize()
        try:
            alert_wait=EC.alert_is_present()
            print ('Alert1')
            print(alert_wait)
            #shell.Sendkeys('{ENTER}')
            wait(driver ,30).until(alert_wait)
            time.sleep(10)
            alert=driver.switch_to_alert()
            print (alert.text)
            #alert.send_keys('{ENTER}')
            print('Enter')
            alert.accept()
        except:
            pass
        aw=True
        while aw:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.Sendkeys('n9941391')#n9941391#n9930786
            shell.Sendkeys('{TAB}')
            print(' user name done')
            shell.Sendkeys('Sep2018$')#Jan2018$#MssCSO001@
            shell.Sendkeys('{ENTER}')
            print('password done')
            aw=False        
        try:
            
            WebDriverWait(driver, 120).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
            print("alert accepted")
        except TimeoutException:
            print("no alert")
    except Exception as e:
        print (e)
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,30).until(alert_wait)
        print ('Alert3')
        print(alert_wait)
        alert=driver.switch_to.alert.accept()
        print (alert.text)
        alert.accept()
    except:
        pass
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,200).until(alert_wait)
        print ('Alert4')
        print(alert_wait)
        alert=driver.switch_to.alert.accept()
        print (alert.text)
        alert.send_keys('{ENTER}')
        print('Enter')
        alert.accept()
    except:
        pass
download_file()
