from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import xlrd
import os
def attain(user_id,password):
    # your executable path is wherever you saved the chrome webdriver
    chromedriver = "C:\\Users\\Praveen\\Downloads\\Compressed\\chromedriver_win32\\chromedriver.exe"
    browser = webdriver.Chrome(executable_path=chromedriver)
    browser.get("http://www.nietcampus.com/Home/")
    email = browser.find_element_by_id("txtUserName")
    email.send_keys(user_id)
    pwd = browser.find_element_by_id("txtPassword")
    pwd.send_keys(password)
    submission = browser.find_element_by_name("Ulogin")
    submission.click()
    browser.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
    browser.get("http://www.nietcampus.com/manage/AttendanceReport/SubjectCodeWiseSemAttendence.aspx")
    time.sleep(2)
    select1= Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlSession"))
    select1.select_by_value("Sess000011")
    time.sleep(2)
    select2=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_DDLInstitute"))
    select2.select_by_value("IM00000001")
    time.sleep(2)
    select3=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlprogram"))
    select3.select_by_value("PG00000001")
    time.sleep(2)
    if(user_id[4:7]=="ITE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR04")
    if(user_id[4:7]=="CIV"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR02")
    if(user_id[4:7]=="CSE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR01")
    if(user_id[4:7]=="MEC"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR03")
    if(user_id[4:7]=="ECE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR05")
    if(user_id[4:7]=="BTE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR08")
    if(user_id[4:7]=="EEE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR06")
    if(user_id[4:7]=="CHE"):
        select4=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlbranch"))
        select4.select_by_value("BR07")
    
    time.sleep(2)
    select5=Select(browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_ddlsem"))
    select5.select_by_value("CS04")
    time.sleep(2)
    view=browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_btnView")
    view.click()
    time.sleep(15)
    xceldown=browser.find_element_by_id("ctl00_ctl00_ContentPane_ContentPane_AttendenceColor_imgexcel")
    xceldown.click()
    time.sleep(10)
    book=xlrd.open_workbook("C:\\Users\\Praveen\\Downloads\\ULTRAWEBGRID1.XLS")
    sheet=book.sheet_by_index(0)
    for rowidx in range(sheet.nrows):
        value=str(sheet.cell(rowidx,1))
        value=value[6:16]
        if(value==user_id):
            break;
    cell=sheet.cell(rowidx,18)
    print(cell)
    percent=str(cell)
    percent=str(percent[7:])
    percent=float(percent)
    cell1 = sheet.cell(rowidx,15)
    cell2=str(cell1)
    no_of_days=str(cell2[10:13])
    no_of_days=float(no_of_days)
    minimum=(((percent/80)-1)*(no_of_days))
    print("No. of lectures u can skip to maintain minimum 80% :" ,int(minimum))
    time.sleep(1)
    os.remove("C:\\Users\\Praveen\\Downloads\\ULTRAWEBGRID1.XLS")
