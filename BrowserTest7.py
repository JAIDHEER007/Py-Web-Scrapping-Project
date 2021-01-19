from selenium import webdriver 
from selenium.webdriver.chrome.options import Options 
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import JavascriptException
 
import platform
from os import getcwd

import re

import csv

csv_file = open('BrowserTest7.csv','w')
csvfile_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)

csvfile_writer.writerow(['Name','Roll Number','Sem1 CGPA'])

opt = Options() 
opt.add_experimental_option("excludeSwitches", ["enable-logging"]) 

prefs = {"profile.managed_default_content_settings.images": 2}
opt.add_experimental_option("prefs", prefs)

opt.add_argument('--ignore-ssl-errors=yes')
opt.add_argument('--ignore-certificate-errors')

#opt.add_argument('--headless')

path = getcwd() + '/chromedriver.exe'

driver = webdriver.Chrome(options = opt, executable_path = path) 

driver.get('https://examsection.aec.edu.in/Login.aspx')



for i in range(2,66):
	if(i == 12): continue
		
	student = "19A91A05"
	
	if((i>=2) and (i<=9)): student += ("0" + str(i))
	else: student += str(i)

	data = list()
	try:
		driver.execute_script("javascript:__doPostBack('lnkStudent','')")
	except JavascriptException:
		print('Could Not Run the Student Login Java Script')
		driver.quit()
		exit()

	try:
		driver.find_element_by_xpath('//*[@id="txtUserId"]').send_keys(student)
		driver.find_element_by_xpath('//*[@id="txtPwd"]').send_keys(student)
		driver.find_element_by_xpath('//*[@id="btnLogin"]').click()
	
		try:
			inc_pwd = driver.find_element_by_xpath('//*[@id="lblWarning"]')
			print(inc_pwd.text)
			driver.quit()
			exit()
		except NoSuchElementException:	
			Ht_No = driver.find_element_by_xpath('//*[@id="lblHTNo"]')
			Name = driver.find_element_by_xpath('//*[@id="lblStudName"]')
	except NoSuchElementException:
		print('Could Not find the given element by XPath')
		driver.quit()
		exit()


	print('Name:',Name.text)
	print('Hall Ticket Number:',Ht_No.text)
	data.append(Name.text)
	data.append(Ht_No.text)
	try:
		driver.execute_script("javascript:__doPostBack('ctl00$lnkOverallMarksSemwise','')")
	except JavascriptException:
		print("Cannot run JS")
		driver.quit()
		exit()

	try:
		cgpa = driver.find_element_by_xpath('//*[@id="cpStudCorner_lblFinalCGPA"]')
		print(cgpa.text)
		
		cgpa_data = re.findall('[0-9]{1}\.?[0-9]*',cgpa.text)
		
		if len(cgpa_data) == 0:
			data.append('N/A')
		else:
			data.append(float(cgpa_data[0]))
		
	except NoSuchElementException:
		print('Could Not find the marks')
		
	
	csvfile_writer.writerow(data)
	
	try:
		log_out = driver.find_element_by_xpath('//*[@id="btnLogout"]')
		try:
			driver.execute_script("arguments[0].click();", log_out)
		except:
			print('Could not run the Logout JS')
			driver.quit()
			exit()
	except NoSuchElementException:
		print("Could not find the Log Out Button")
 
yn = input('Do you want to close (Y/N): ')[0]
if yn in 'Yy': 
	csv_file.close()
	driver.quit() #Used to close the driver
	
