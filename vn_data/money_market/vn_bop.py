from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from pandas.tseries.offsets import MonthEnd

import time, datetime

year = int(input('Year: '))
quarter = int(input('Quarter: '))

quarter_dict = {'3':'I', '6':'II', '9':'III', '12':'IV'}

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

def bop(date):
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
	driver.get("https://sbv.gov.vn/webcenter/portal/vi/menu/trangchu/tk/ccttqt")
	driver.find_element(By.NAME,'T:oc_7650641436region:id1').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.NAME,'T:oc_7650641436region:id4').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.ID,'T:oc_7650641436region:cb1').click()
	time.sleep(3)
	element = driver.find_element(By.ID, 'T:oc_7650641436region:t1').text
	quarter_condition = quarter_dict[datetime.datetime.strftime(date, '%#m')]+'/'+datetime.datetime.strftime(date, '%Y')
	# print(quarter_condition)
	if quarter_condition in element:
		driver.find_element(By.ID,'T:oc_7650641436region:t1:0:cl3').click()
		time.sleep(30)

bop(datetime.datetime(year, quarter*3, 1) + MonthEnd(0))