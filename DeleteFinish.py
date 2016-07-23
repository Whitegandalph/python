#!/usr/bin/python2
import openpyxl
import os
import webbrowser
import time
import shutil
import smtplib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
#from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0


path = '/home/pi/working/'
os.chdir(path)
filelist = os.listdir(path)[0]

wb = openpyxl.load_workbook(filelist)
names =wb.get_sheet_names()


if "DeleteFinish" in names:
    print "DeleteFinish"
else:
    print "exiting"
    exit()



sheet = wb.get_sheet_by_name('DeleteFinish')
user = sheet['B1'].value
code= sheet['B2'].value

browser=webdriver.Firefox()
browser.get('https://work.fleetmatics.com/')
emailElem = browser.find_element_by_id('USER_NAME')
emailElem.send_keys(user)
passwordElem = browser.find_element_by_id('PASSWORD')
passwordElem.send_keys('1234')
button = browser.find_element_by_id('ImageButton1').click()

browser.get('https://work.fleetmatics.com/Admin/Options/JobFinishQuestionList.aspx')

for x in range(0,3000):
    try:
        elem = browser.find_element_by_id('ctl00_BodyContentPlaceHolder_GridView1_ctl02_ibtnDelete')
        elem.click()
        print "Successful"
        browser.implicitly_wait(1)
    except:
        print "Element is not visible"
        filelist = os.listdir(path)[0]
        asfile = '/home/pi/done/' + time.ctime() + filelist + '.xlsx'
        shutil.move(filelist, asfile)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login("david.balla.fleetmatics@gmail.com", "Fleet2015") 
        msg = "DeleteQuestions Complete " + asfile
        server.sendmail("david.balla.fleetmatics@gmail.com", "davidballa@gmail.com", msg)
        server.quit()
        browser.get('https://work.fleetmatics.com/Admin/Logout.aspx')
        print "done"
        browser.implicitly_wait(3)
        browser.close()

        
    
