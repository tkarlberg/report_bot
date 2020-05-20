from selenium import webdriver
from time import sleep
import pandas as pd
from secrets import user,pw,to_path,save_path
import datetime
import calendar
import os

# BOT ACTIONS
# load https://tms.ezfacility.com/reports/report_attendance_overview.aspx
# input email and password
# click login
# change start and end date to month begin and end
# cycle through trainer with in each cycle through package types
# Generate report
# Save report to csv (name: trainer month package type)
# 
# GENERATE REPORT
# Cycle through each trainer for following:
# Read in all reports
# get sum of vists for each report
# make array of trainer month pyt ptpkg performance etc
# append array to csv

class report_bot(): 
    def __init__(self):
        self.driver = webdriver.Chrome()
    def login(self, name, pswd):
        self.driver.get('https://tms.ezfacility.com/reports/report_attendance_overview.aspx')

        sleep(2)

        username = self.driver.find_element_by_xpath('/html/body/div[1]/form/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div/input')
        username.send_keys(name)

        password = self.driver.find_element_by_xpath('/html/body/div[1]/form/div[3]/div[1]/div[1]/div[2]/div[2]/div[2]/div/input')
        password.send_keys(pswd)
        
        login_button = self.driver.find_element_by_xpath('/html/body/div[1]/form/div[3]/div[1]/div[1]/div[2]/div[2]/div[4]/input')
        login_button.click()
        sleep(5)

    def set_dates(self):
        start = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_DatePickerStart_txtTextBox"]')
        date1 = pd.to_datetime(start.get_attribute('defaultValue'))
        d2y, d2m = calendar.prevmonth(date1.year,date1.month)
        date2 = datetime.date(d2y,d2m,1)
        startdate = str(date2.month) +'/'+ str(date2.day) +'/'+ str(date2.year)
        start.clear()
        start.send_keys(startdate)
        
        end = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_DatePickerEnd_txtTextBox"]')
        enddate = str(date2.month) + '/' + str(calendar.monthrange(date2.year,date2.month)[1]) + '/' + str(date2.year)
        end.clear()
        end.send_keys(enddate)
        return date2
    
    def get_trainers(self):
        trainerlist = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_drpTrainer"]').get_attribute('textContent')
        trainers = trainerlist.split(sep = '\n')[2:-2]
        return trainers
    
    def get_package_types(self):
        packagelist = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_drpPackageType"]').get_attribute('textContent')
        packages = packagelist.split(sep = '\n')[2:-1]
        return packages
    
    def generate_report(self, trainer, package):
        set_trainer = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_drpTrainer"]')
        set_trainer.send_keys(trainer)
        sleep(2)
    
        set_package = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_drpPackageType"]')
        set_package.send_keys(package)
    
        generate_button = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_btnGenerateReport"]')
        generate_button.click()
        sleep(2)

    def download_report(self, trainer, package):
        dropdown = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_rv_ctl09_ctl04_ctl00_ButtonImgDown"]')
        dropdown.click()
        csvbutton = self.driver.find_element_by_xpath('//*[@id="ctl00_ctl00_ctl00_ctl00_CphFormBody_cphFormBody_cphRight_cphReport_rv_ctl09_ctl04_ctl00_Menu"]/div[7]/a')
        csvbutton.click()
        sleep(4)
        os.rename('C:/Users/Tyler/Downloads/Attendance_Overview.csv',  to_path + trainer + package + '.csv')

bot = report_bot()
bot.login(user, pw)
date = bot.set_dates()
month = date.month
year = date.year
trainers = bot.get_trainers()
packages = bot.get_package_types()
packages.remove('Other Duties')
packages.remove('Unavailable')
packages.remove('Isolated')
packages.remove('Park Workout')

packages = ['PYT', 'PT Package']

for trainer in trainers:
    for package in packages:
        bot.generate_report(trainer,package)
        bot.download_report(trainer,package)

days = ['Monday', 'Tuesday', 'Wednesday','Thursday','Friday','Saturday','Sunday']
headers = ['Trainer']+packages
report = pd.DataFrame(columns = headers)
for trainer in trainers:
    totals = [trainer]
    for package in packages:
        total = 0
        try:
            temp = pd.read_csv(to_path + '/'+ trainer + package + '.csv', nrows = 11)
            temp = temp.drop(columns = 'StartTimeDate2')
            temp = temp.loc[temp['textbox10'].isin(days)]
            temp['txtScheduleGroup'] = temp['txtScheduleGroup'].astype('int32')
            total += temp.sum(axis=0).iloc[1]
        except:
            total = 0
        totals.append(total)
    report = report.append(pd.DataFrame([totals], columns = headers))
report['PT Total'] = report['PYT'] + report['PT Package']
report.reset_index(drop = True, inplace = True)
report.loc[len(trainers),:] = report.sum()
report.iloc[-1,0] = 'Totals'
report.to_csv(save_path + calendar.month_name[month] + str(date.year) + '.csv', index = False)

files = os.listdir(to_path)
for file in files:
    os.remove(to_path + file)

bot.driver.quit()