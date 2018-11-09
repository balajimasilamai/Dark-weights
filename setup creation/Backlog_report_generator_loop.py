
from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.select import Select
#import os
import win32com.client

#import datetime

def download_file(query,file_name):
    file_names=file_name
    ##################################################
    #This portion is added in UI.py in order to avoid multiple importing of the same library module
    '''
    #To get the current date and time
    file_names=['Open_',
            'Closed_',
            'Request_Open_',
            'Paetec_',
            'Paetec_Closed_',
            'paetec_request_',
            'paetec_ris_'
            ]
    current_time=datetime.date.today().strftime("%B %dth, %Y")
    a=datetime.date.today().strftime("%B")
    b=datetime.date.today().strftime("%dth")
    c=datetime.date.today().strftime("%Y")
    print (current_time)
    print (a)
    print (b)
    print (c)
    file_name=b[0:]+a[0:3]+c[2:4]
    print (file_name)

    for i in range(0,len(file_names)):
        print (file_names[i]+str(file_name))
    '''

    ###########################################
    
    #Selenium part to download the files
    #with VPN
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
            wait(driver ,30).until(alert_wait)
            alert=driver.switch_to_alert()
            print (alert.text)
            alert.accept()
        except:
            pass
        aw=True
        while aw:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.Sendkeys('n9941391')#n9941391#n9930786
            shell.Sendkeys('{TAB}')
            shell.Sendkeys('Sep2018$')#Jan2018$#MssCSO001@
            shell.Sendkeys('{ENTER}')
            aw=False
            try:
                alert_wait=EC.alert_is_present()
                wait(driver ,120).until(alert_wait)
                alert=driver.switch_to_alert()
                #alert=driver.switch_to.alert.accept()
                print (alert.text)
                alert.send_keys('{ENTER}')
                alert.accept()
            except:
                pass
        
    except Exception as e:
        print (e)
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,30).until(alert_wait)
        alert=driver.switch_to_alert()
        print (alert.text)
        alert.accept()
    except:
        pass
    try:
        alert_wait=EC.alert_is_present()
        wait(driver ,200).until(alert_wait)
        alert=driver.switch_to.alert.accept()
        print (alert.text)
        alert.accept()
    except:
        pass
    time.sleep(10)
    appclicker=driver.find_element_by_xpath('//*[@id="reg_img_304316340"]').click()
    time.sleep(5)

    print ('Application button clicked')

    element=driver.find_element_by_xpath("//span[text()='Incident Management']") 

    webdriver.ActionChains(driver).move_to_element(element).click(element).perform() 
    time.sleep(10)

    driver.find_element_by_xpath("//span[text()='Search Incident']").click()
    #alerter()
    print ('Going to Search Indicent')
    time.sleep(15)
    # Clicking Adnavce search buton
    driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[1]/table/tbody/tr/td[3]/a[3]").click()
    
    for i in range(0,len(query)):
        #print (query[i])
        body = driver.find_element_by_css_selector('body')
        body.send_keys(Keys.PAGE_DOWN)
        if i!=0:
            time.sleep(6)
            #Xpath to select New search
            new_search=driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[5]/div[2]/div[1]/div[1]/div[3]/fieldset[1]/div[1]/div[1]/div[1]/div[1]/div[3]/fieldset[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/a[1]/span[1]")
            new_search.click()
        # Clicking Text Area
        time.sleep(0.025)
        text_area=driver.find_element_by_xpath('//fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[5]/table[2]/tbody/tr/td[1]/textarea[@id="arid1005"]')
        time.sleep(0.025)
        text_area.click()
        time.sleep(0.025)
        text_area.send_keys(query[i])
        time.sleep(15)
        print ('Searching')
        # Clicking the search button
        driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[4]/div[4]/div/div/div[3]/fieldset/div/div/div/div/div[2]/fieldset/div/a[2]/div/div').click() #search
        try:
            alert_wait=EC.alert_is_present()
            wait(driver ,40).until(alert_wait)
            alert=driver.switch_to_alert()
            print (alert.text)
            alert.accept()
        except:
            pass
        time.sleep(30)
        #text_area.clear()
        #selecting all button
        driver.find_element_by_xpath('html[1]/body[1]/div[1]/div[5]/div[2]/div[1]/div[1]/div[3]/fieldset[1]/div[1]/div[1]/div[1]/div[1]/div[3]/fieldset[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/a[2]').click()  #selectall
        time.sleep(2)
        #Selecting Report button
        driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[2]/div/div/div[3]/fieldset/div/div/div/div/div[3]/fieldset/div/div/div/div[2]/div/div[3]/table/tbody/tr/td[1]/a[1]').click() #report
        # New Window will open
        window_after = driver.window_handles[1]
        time.sleep(15)
        #Trnasfreing the window control
        driver.switch_to.window(window_after)
        time.sleep(15)
        #searching the Prodapt_ASAP_Daily_Report and clicking
        driver.find_element_by_xpath("//span[text()='Prodapt_ASAP_Daily_Report']").click()
        s=(driver.find_element_by_xpath("//*[@id='arid_WIN_0_2000053']")).click() #To clcick the Destination drop down
        time.sleep(0.025)
        f=driver.find_element_by_xpath('/html/body/div[3]/div[2]/table/tbody/tr[2]/td[1]')# to select the File
        webdriver.ActionChains(driver).move_to_element(f).click(f).perform()
        time.sleep(0.5)
        w=driver.find_element_by_xpath("//*[@id='arid_WIN_0_2000056']").click()#Formtdropdownbutton
        c=driver.find_element_by_xpath('/html/body/div[3]/div[2]/table/tbody/tr[2]/td[1]')# clicking CSV
        webdriver.ActionChains(driver).move_to_element(c).click(c).perform()
        time.sleep(0.5)
        #fname=driver.find_element_by_xpath('//*[@id="WIN_0_2000057"]/a/img').click() #to click the Filename menu button
        textname=driver.find_element_by_xpath("//*[@id='arid_WIN_0_2000057']")#to click the filename
        time.sleep(0.002)
        textname.click()
        textname.clear()
        textname.send_keys((file_names[i]+'.csv'))
        time.sleep(0.5)
        driver.find_element_by_xpath("//*[@id=\"reg_img_93272\"]").click()
        print ('file is downloaded')
        time.sleep(25)
        driver.close()
        driver.switch_to.window(window_before)
    driver.close()    
