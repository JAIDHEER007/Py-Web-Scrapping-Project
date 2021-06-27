from selenium import webdriver 
from selenium.webdriver.chrome.options import Options 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import JavascriptException
from os import getcwd
from time import perf_counter
import re
import csv
from datetime import datetime

time_stamp = str(datetime.now()).replace(':','-')

#Working with CSV File
csv_file = open('BrowserTest13 '+time_stamp+'.csv','w')
csvfile_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
# csvfile_writer.writerow(['Name','Roll Number','Sem1 SGPA','Sem2 SGPA','Sem3 SGPA','CGPA','Time Taken (sec)'])
csvfile_writer.writerow(['Name','Roll Number','Sem3 SGPA','CGPA','Time Taken (sec)'])


#Disabling console logging
opt = Options() 
opt.add_experimental_option("excludeSwitches", ["enable-logging"]) 

#Stopping images from Loading 
prefs = {"profile.managed_default_content_settings.images": 2}
opt.add_experimental_option("prefs", prefs)

#Avoiding "Your Connection is not private" error
opt.add_argument('--ignore-ssl-errors=yes')
opt.add_argument('--ignore-certificate-errors')

#opt.add_argument('--headless')

path = 'E:\COMPUTER\Python 3\experiments\Chrome Driver\chromedriver.exe'

driver = webdriver.Chrome(options = opt, executable_path = path) 

driver.get('https://examsection.aec.edu.in/Login.aspx')

le_students = ['20A95A0501', '20A95A0502', '20A95A0503', '20A95A0504', '20A95A0505', '20A95A0506', '20A95A0507']

for student in le_students:
    start = perf_counter()
    # if(i == 12): continue
    
    # student = "19A91A05"
    
    # if((i>=2) and (i<=9)): student += ("0" + str(i))
    # else: student += str(i)

    data = list()
    try:
        driver.execute_script("javascript:__doPostBack('lnkStudent','')")
        driver.find_element_by_id('txtUserId').send_keys(student)
        driver.find_element_by_id('txtPwd').send_keys(student)
        driver.find_element_by_id('btnLogin').click()
    
        Ht_No = driver.find_element_by_id('lblHTNo')
        Name = driver.find_element_by_id('lblStudName')
        print('Name:',Name.text)
        print('Hall Ticket Number:',Ht_No.text)
        data.append(Name.text)
        data.append(Ht_No.text)
        driver.execute_script("javascript:__doPostBack('ctl00$lnkOverallMarks','')")

        # sem1 = driver.find_element_by_xpath('//*[@id="cpStudCorner_pnMarks"]/table[1]/tbody/tr[13]/td[3]')
        # print(sem1.text)
    
        # sem1SGPA = re.findall('[0-9]{1}\.?[0-9]*',sem1.text)[0]
        # data.append(sem1SGPA)
    
        # sem2 = driver.find_element_by_xpath('//*[@id="cpStudCorner_pnMarks"]/table[2]/tbody/tr[13]/td[3]')
        # print(sem2.text)
    
        # sem2SGPA = re.findall('[0-9]{1}\.?[0-9]*',sem2.text)[0]
        # data.append(sem2SGPA)

        # sem3 = driver.find_element_by_xpath('//*[@id="cpStudCorner_pnMarks"]/table[3]/tbody/tr[13]/td[3]')
        # print(sem3.text)
    
        # sem3res = re.findall('[0-9]{1}\.?[0-9]*',sem3.text)
        # data.append(sem3res[0])
        # data.append(sem3res[1])
        
        sem = driver.find_element_by_xpath('//*[@id="cpStudCorner_pnMarks"]/table/tbody/tr[13]/td[3]')
        print(sem.text)
        
        sem_res = re.findall('[0-9]{1}\.?[0-9]*',sem.text)
        
        data.append(sem_res[0])
        data.append(sem_res[1])
        
        log_out = driver.find_element_by_id('btnLogout')
        driver.execute_script("arguments[0].click();", log_out)
    finally:
        data.append(perf_counter()-start)
        csvfile_writer.writerow(data)
    
yn = input('Do you want to close (Y/N): ')[0]
if yn in 'Yy': 
    csv_file.close()
    driver.quit() #Used to close the driver
    
