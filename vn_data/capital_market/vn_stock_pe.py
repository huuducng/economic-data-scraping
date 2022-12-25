import requests, datetime, openpyxl, pathlib

home = str(pathlib.Path.home())+'/CloudDrive/My Drive/data/vn_stock_market/'

# load data
headers = {
'Accept':'application/json, text/plain, */*',
'Accept-Language':'en-US,en;q=0.9',
'Authorization':'Bearer',
'Connection':'keep-alive',
'DNT':'1',
'Origin':'https://fiintrade.vn',
'Referer':'https://fiintrade.vn/',
'Sec-Fetch-Dest':'empty',
'Sec-Fetch-Mode':'cors',
'Sec-Fetch-Site':'same-site',
'Sec-GPC':'1',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.67 Safari/537.36',
}
params = {
'language':'vi',
'Code':'VNINDEX',
'TimeRange':'OneYear',
'FromDate':'',
'ToDate':'',
}
url = 'https://market.fiintrade.vn/MarketInDepth/GetValuationSeriesV2'
response = requests.get(url=url,headers=headers,params=params).json()
data = response['items']
oldest = datetime.datetime.strptime(data[0]['tradingDate'][:10],'%Y-%m-%d')

wb = openpyxl.load_workbook(home+'vn_pe.xlsx')
ws = wb.active

for row in ws.rows:
	if row[0].value == oldest:
		break

p = int(row[0].row)

for r in range(len(data)):
	ws.cell(row=r+p,column = 1,value=datetime.datetime.strptime(data[r]['tradingDate'][:10],'%Y-%m-%d').date())
	ws.cell(row=r+p,column = 2,value=data[r]['r21'])
	ws.cell(row=r+p,column = 3,value=data[r]['r25'])

wb.save(home+'vn_pe.xlsx')