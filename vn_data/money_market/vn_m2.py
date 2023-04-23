from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from pandas.tseries.offsets import MonthEnd

import time, datetime

year = int(input('Year: '))
month = int(input('Month: '))

def tableDataText(table):    
    """Parses a html segment started with tag <table> followed 
    by multiple <tr> (table rows) and inner <td> (table data) tags. 
    It returns a list of rows with inner columns. 
    Accepts only one <th> (table header/data) in the first row.
    """
    def rowgetDataText(tr, coltag='td'): # td (data) or th (header)       
        return [td.get_text(strip=True) for td in tr.find_all(coltag)]  
    rows = []
    trs = table.find_all('tr')
    headerow = rowgetDataText(trs[0], 'th')
    if headerow: # if there is a header row include first
        rows.append(headerow)
        trs = trs[1:]
    for tr in trs: # for every table row
        rows.append(rowgetDataText(tr, 'td') ) # data row       
    return rows

def m2(date):
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
	driver.get("https://sbv.gov.vn/webcenter/portal/vi/menu/trangchu/tk/pttt/tpttt")
	driver.find_element(By.NAME,'T:oc_0810304495region:id1').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.NAME,'T:oc_0810304495region:id4').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.ID,'T:oc_0810304495region:cb1').click()
	time.sleep(3)
	element = driver.find_element(By.ID, 'T:oc_0810304495region:t1').text
	month_condition = datetime.datetime.strftime(date, '%#m')+'/'+datetime.datetime.strftime(date, '%Y')
	# print(quarter_condition)
	if month_condition in element:
		driver.find_element(By.ID,'T:oc_0810304495region:t1:0:cl3').click()
		time.sleep(30)

m2(datetime.datetime(year, month, 1) + MonthEnd(0))