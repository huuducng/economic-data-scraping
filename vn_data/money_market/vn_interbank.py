from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time, datetime, openpyxl, pathlib

engsub = {"Quađêm":"ON","1Tuần":"1W","2Tuần":"2W","1Tháng":"1M","3Tháng":"3M","6Tháng":"6M","9Tháng":"9M","12Tháng":"12M"}

home = str(pathlib.Path.home())+'/data/vn_money_market/'
today = datetime.datetime.today() + datetime.timedelta(-1)

def getlslnh(date):
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
	driver.get("https://www.sbv.gov.vn/webcenter/portal/vi/menu/rm/ls/lsttlnh")
	driver.find_element(By.NAME,'T:oc_5531706273region:id1').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.NAME,'T:oc_5531706273region:id4').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.ID,'T:oc_5531706273region:cb1').click()
	time.sleep(1.5)
	if len(driver.find_elements(By.LINK_TEXT,'Xem')) > 0:
		driver.find_element(By.LINK_TEXT,'Xem').click()
		time.sleep(1.5)
		raw = driver.find_elements(By.CLASS_NAME,'jrPage')[1].text.splitlines()
		result1 = [x for x in raw if 'Qua đêm' in x]
		result2 = [x for x in raw if 'Tuần' in x]
		result3 = [x for x in raw if 'Tháng' in x]
		result_raw = result1+result2+result3
		result = {}
		for item in result_raw:
			sub = []
			subdict = {}
			item = item.split()
			period_V = item[0]+item[1]
			period_E = engsub[period_V]
			if (len(item) <= 4) and ('*' not in item[2]):
				sub.append(float(item[2].replace(',','.').replace('%','')))
			else:
				sub.append('')
			if (len(item) == 4) and ('*' not in item[3]):
				sub.append(float(item[3].replace('.','').replace(',','.')))
			else:
				sub.append('')
			subdict.update({period_E:sub})
			result.update(subdict)

		return result
	driver.quit()

wb = openpyxl.load_workbook(home+'vn_interbank.xlsx')
ws = wb.active

if ws.cell(row=len(ws['A']),column=1).value.weekday() == 4:
	latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(3)
else:
	latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(1)

if latest>today:
	quit()

r=len(ws['A'])+1

while True:
	print(latest)
	if getlslnh(latest):
		a = getlslnh(latest)
		for i in range(len(a)):
			ws.cell(row=r+i, column=1, value=latest.date())
			
		for item in a:
			ws.cell(row=r, column=2, value=item)
			ws.cell(row=r, column=3, value=a[item][0])
			ws.cell(row=r, column=4, value=a[item][1])
			r+=1
	
	if latest.weekday() == 4:
		latest = latest + datetime.timedelta(3)
	else:
		latest = latest + datetime.timedelta(1)

	wb.save(home+'vn_interbank.xlsx')
	if latest > today:
		break