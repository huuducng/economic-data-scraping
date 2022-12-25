from bs4 import BeautifulSoup
import requests, datetime, openpyxl, pathlib

home = str(pathlib.Path.home())+'/CloudDrive/My Drive/data/vn_money_market/'

# functions
def tableDataText(table):
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

def giavang(date):
	response = requests.get(url='https://www.giavangonline.com/goldhistory.php',params={'date':date.strftime('%Y-%m-%d')})
	soup = BeautifulSoup(response.text,'html.parser')
	htmltable = soup.find_all('table')[3]
	newchar1 = ''
	newchar2 = ''
	list_result = []

	if len(tableDataText(htmltable)) > 3:
		list_table = tableDataText(htmltable)[2]
		
		where = list_table[1].index('/')
		for char in list_table[1][:where-1]:
			newchar1 = newchar1 + char
		newchar1 = int(newchar1.replace(',','')+'0')
		

		for char in list_table[1][where+1:]:
			newchar2 = newchar2 + char
		newchar2 = int(newchar2.replace(',','')+'0')
	list_result.append(date.date())
	list_result.append(newchar1)
	list_result.append(newchar2)
	return list_result

end_date = datetime.date.today()

wb = openpyxl.load_workbook(home+'vn_sjcgold.xlsx')
ws = wb.active

r = len(ws["A"])

date = ws.cell(row=r, column=1).value

while True:
	if date.weekday() == 4:
		date = date + datetime.timedelta(3)
	else:
		date = date + datetime.timedelta(1)
	if date.date() >= datetime.datetime.today().date():
		break
	r+=1
	print(date.strftime('%d/%m/%Y'))
	for item in giavang(date):
		ws.cell(row=r,column=giavang(date).index(item)+1,value=item)


wb.save(home+'vn_sjcgold.xlsx')