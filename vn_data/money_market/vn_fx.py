from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

from bs4 import BeautifulSoup
import time, datetime, openpyxl, requests, pathlib
import pandas as pd

home = str(pathlib.Path.home())+'/data/vn_fx/'
today = datetime.datetime.today() + datetime.timedelta(-1)

def tudo(date):
	response = requests.get(url='https://tygiadola.net/TyGia',params={'date':datetime.datetime.strftime(date,'%d-%m-%Y')})
	soup = BeautifulSoup(response.text, 'html.parser')
	htmltables = soup.find_all('table')
	result_list = []
	inter = []
	for table in htmltables:
		for c in table.find_all('th'):
			if 'tự do' in c.get_text():
				x = table.find_all('td',{'class':'text-right'})
				for item in x:
					inter.append(item.get_text())
	for item in inter:
		result_list.append(int(item.split(' ')[0].replace(',','')))

	return result_list

def sbv(date):
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
	driver.get("https://www.sbv.gov.vn/TyGia/faces/TyGiaTrungTam.jspx")
	driver.find_element(By.NAME,'pt1:r2:0:id1').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.NAME,'pt1:r2:0:id4').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.ID,'pt1:r2:0:cb1').click()
	time.sleep(1.5)
	if len(driver.find_elements(By.LINK_TEXT,'Xem')) > 0:
		driver.find_element(By.LINK_TEXT,'Xem').click()
		time.sleep(1)
		raw = driver.find_elements(By.CLASS_NAME,'jrPage')[1].text.splitlines()

		if raw[-1].split()[-1] == datetime.datetime.strftime(date, '%d/%m/%Y'):
			return {date:int(raw[2].split()[-2].replace('.',''))}

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

def vcb(date):
	response = requests.get(url='https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx',params={'txttungay':date.strftime('%d/%m/%Y')})
	soup = BeautifulSoup(response.text, 'html.parser')
	htmltable = soup.find_all('table')[0]
	list_table = tableDataText(htmltable)
	for row in list_table:
		if len(row) == 5:
			if row[1] == 'USD':
				return [float(row[2].replace(',','')),float(row[3].replace(',','')),float(row[4].replace(',',''))]


wb = openpyxl.load_workbook(home+'vn_fx.xlsx')
ws = wb.active

if ws.cell(row=len(ws['A']),column=1).value.weekday() == 4:
	latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(3)
else:
	latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(1)

if latest>today:
	quit()

r=len(ws['A'])+1

while True:
	if sbv(latest):
		ws.cell(row=r, column=1, value=latest.date())
		ws.cell(row=r, column=2, value=sbv(latest)[latest])
		if tudo(latest):
			ws.cell(row=r, column=3, value=tudo(latest)[0])
			ws.cell(row=r, column=4, value=tudo(latest)[1])
		if vcb(latest):
			ws.cell(row=r, column=5, value=vcb(latest)[0])
			ws.cell(row=r, column=6, value=vcb(latest)[1])
			ws.cell(row=r, column=7, value=vcb(latest)[2])
		r+=1
	if latest.weekday() == 4:
		latest = latest + datetime.timedelta(3)
	else:
		latest = latest + datetime.timedelta(1)

	if latest > today:
		break

wb.save(home+'vn_fx.xlsx')

fx = pd.read_excel(home+'vn_fx.xlsx', sheet_name=0)
fx.set_index('date').resample('M').mean().reset_index().to_excel(home+'vn_fxM.xlsx', index=False)
fx.set_index('date').resample('Q').mean().reset_index().to_excel(home+'vn_fxQ.xlsx', index=False)