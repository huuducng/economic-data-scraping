import requests, datetime, openpyxl, pathlib
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta

# Turn off warnings
import warnings
warnings.filterwarnings("ignore")

home = str(pathlib.Path.home())+"/CloudDrive/My Drive/data/vn_bond_market/"

# Function
def tableDataText(table):    
    def rowgetDataText(tr, coltag="td"): # td (data) or th (header)       
        return [td.get_text(strip=True) for td in tr.find_all(coltag)]  
    rows = []
    trs = table.find_all("tr")
    headerow = rowgetDataText(trs[0], "th")
    if headerow: # if there is a header row include first
        rows.append(headerow)
        trs = trs[1:]
    for tr in trs: # for every table row
        rows.append(rowgetDataText(tr, "td") ) # data row       
    return rows

# Today
today = datetime.datetime.today()

# Main
def bidding_result(start_date):
    end_date = start_date + relativedelta(days=-1, months=1)
    headers = {
        "Cookie":"language=vi-VN; 616a3745ee32423b8ef6bed543a12282=pjvbk3cjdjcvmcgbw3pyyft3; __RequestVerificationToken_4TPT=yZ4BTA5S0-nn2RT9obdHgoVMUgy-5XTEJKGLRf43ruHMhQ8abgO1Y6vZFHYkoG74igKq2FsRXd-MnYZZc_IAidpMIQM7DMvBowyjPV6zfZ41",
        "X-Requested-With":"XMLHttpRequest",
        "__RequestVerificationToken":"wJ888Z3xodgSNUfQxhbWMFHkmGs0Qg58d7p44donVlkce34hz1lLAnWiigGChMCFNHEgH7mj5_RsX1j3JPfQQJHCD9hgSu7Wjb9MQAYlgAo1",
    }

    a = datetime.datetime.strftime(start_date,"%d/%m/%Y")+"|"+datetime.datetime.strftime(end_date,"%d/%m/%Y")+"|0||0|'VND'|0|1"

    postdata1 = {
        "p_keysearch":a,
        "pColOrder":"col_x",
        "pOrderType":"ASC",
        "pCurrentPage":"1",
        "pRecordOnPage":"10",
        "pIsSearch":"1",
    }

    response = requests.post(url="https://hnx.vn/ModuleReportBonds/Bond_DauThau/Bond_ThongKe_DauThau", headers=headers, data=postdata1, verify=False)
    soup = BeautifulSoup(response.text, "html.parser")
    htmltable = soup.find_all("table")[0]
    list_table = tableDataText(htmltable)[1:-1]

    result = []
    for row in list_table:
        row = row[:7]
        row[0] = end_date.date()
        for i in [3, 4, 5, 6]:
            row[i] = int(row[i].replace(".",""))
        result.append(row)
    return result

wb = openpyxl.load_workbook(home+"vn_tpcp_bidding.xlsx")
ws = wb.active

latest = ws.cell(row=len(ws['A']),column=1).value
start_date = datetime.datetime(latest.year, latest.month, 1) + relativedelta(months=1)

r=len(ws['A'])+1

while True:
	if (start_date + relativedelta(months=2)) > today:
		break
	a = bidding_result(start_date)
	for row in a:
		c=1
		for col in row:
			ws.cell(row=r, column=c, value=col)
			c+=1
		r+=1
	start_date = start_date + relativedelta(months=1)
	if (start_date + relativedelta(months=1)) > today:
		break

wb.save(home+"vn_tpcp_bidding.xlsx")

# Government bond yield
wb = openpyxl.load_workbook(home+'vn_tpcp_yield.xlsx')
ws = wb.active
latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(1)
r=len(ws['A'])+1

while True:
	data = {
	'p_keysearch':latest.strftime('%d/%m/%Y')+'|',
	'pColOrder':'col_a',
	'pOrderType':'ASC',
	'pCurrentPage':'1',
	'pIsSearch':'1',}
	headers = {
        "Cookie":"language=vi-VN; 616a3745ee32423b8ef6bed543a12282=pjvbk3cjdjcvmcgbw3pyyft3; __RequestVerificationToken_4TPT=yZ4BTA5S0-nn2RT9obdHgoVMUgy-5XTEJKGLRf43ruHMhQ8abgO1Y6vZFHYkoG74igKq2FsRXd-MnYZZc_IAidpMIQM7DMvBowyjPV6zfZ41",
        "X-Requested-With":"XMLHttpRequest",
        "__RequestVerificationToken":"wJ888Z3xodgSNUfQxhbWMFHkmGs0Qg58d7p44donVlkce34hz1lLAnWiigGChMCFNHEgH7mj5_RsX1j3JPfQQJHCD9hgSu7Wjb9MQAYlgAo1",
    }

	response = requests.post(url='https://hnx.vn/ModuleReportBonds/Bond_YieldCurve/SearchAndNextPageDuLieuTT_Indicator', headers=headers, data=data, verify=False)
	soup = BeautifulSoup(response.text, 'html.parser')
	htmltable = soup.find_all('table')[0]
	list_table = tableDataText(htmltable)[1:]
	for row in list_table:
		c=2
		for column in row:
			if c > 2 and len(column) > 0:
				ws.cell(row=r,column=c,value=float(column.replace(',','.')))
			else:
				ws.cell(row=r,column=c,value=column)
			c+=1
		ws.cell(row=r,column=1,value=latest.date())
		r+=1
	if latest.weekday() == 4:
		latest = latest + datetime.timedelta(3)
	else:
		latest = latest + datetime.timedelta(1)
	if latest > datetime.datetime.today()+datetime.timedelta(-1):
		break

wb.save(home+'vn_tpcp_yield.xlsx')

# Government bond issuance coupon rate

def get_daily_ir(date):
	headers = {
	'Cookie':'language=vi-VN; 616a3745ee32423b8ef6bed543a12282=pjvbk3cjdjcvmcgbw3pyyft3; __RequestVerificationToken_4TPT=yZ4BTA5S0-nn2RT9obdHgoVMUgy-5XTEJKGLRf43ruHMhQ8abgO1Y6vZFHYkoG74igKq2FsRXd-MnYZZc_IAidpMIQM7DMvBowyjPV6zfZ41',
	'X-Requested-With':'XMLHttpRequest',
	'__RequestVerificationToken':'wJ888Z3xodgSNUfQxhbWMFHkmGs0Qg58d7p44donVlkce34hz1lLAnWiigGChMCFNHEgH7mj5_RsX1j3JPfQQJHCD9hgSu7Wjb9MQAYlgAo1',}
	a = datetime.datetime.strftime(date,'%d/%m/%Y')+'|'+datetime.datetime.strftime(date,'%d/%m/%Y')+"|0||0|'VND'|'KBNN'|0"

	postdata = {
	'p_keysearch':a,
	'pColOrder':'col_x',
	'pOrderType':'ASC',
	'pCurrentPage':'1',
	'pRecordOnPage':'10',
	'pIsSearch':'1',}

	result_dict = {'5 Năm':'','7 Năm':'','10 Năm':'','15 Năm':'','20 Năm':'','30 Năm':'',}
	result_list = [date.date()]

	response = requests.post(url='https://hnx.vn/ModuleReportBonds/Bond_DauThau/Bond_ThongKe_DauThau', headers=headers, data=postdata, verify=False)
	soup = BeautifulSoup(response.text, 'html.parser')
	htmltable = soup.find_all('table')[0]
	list_table = tableDataText(htmltable)
	if len(list_table) >= 2:
		for row in list_table[1:-1]:
			if int(row[-3].replace('.','')) > 0:
				result_dict.update({row[1]:float(row[-1].split('-')[-1].replace(',','.'))})
	
	for item in result_dict:
		result_list.append(result_dict[item])
	return result_list

wb = openpyxl.load_workbook(home+'vn_tpcp_rate.xlsx')
ws = wb.active
latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(1)
r=len(ws['A'])+1

start_date = latest

r=len(ws['A'])+1
while True:
	if start_date + datetime.timedelta(1) > datetime.datetime.today():
		break
	
	row = get_daily_ir(start_date)
	for i in range(len(row)):
		ws.cell(row=r, column=i+1, value=row[i])
	r+=1
	if start_date.weekday()==4:
		start_date = start_date + datetime.timedelta(3)
	else:
		start_date = start_date + datetime.timedelta(1)

wb.save(home+'vn_tpcp_rate.xlsx')