from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time, datetime, openpyxl, pathlib

home = str(pathlib.Path.home())+'/CloudDrive/My Drive/data/vn_money_market/'
today = datetime.datetime.today() + datetime.timedelta(-1)

def getmmdata(date):
	result_table = []
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
	driver.get("https://sbv.gov.vn/webcenter/portal/vi/menu/trangchu//hdtttt")
	driver.find_element(By.NAME,'T:oc_2127165840region:id1').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.NAME,'T:oc_2127165840region:id4').send_keys(datetime.datetime.strftime(date, '%d/%m/%Y'))
	driver.find_element(By.ID,'T:oc_2127165840region:cb1').click()
	time.sleep(1.5)
	if len(driver.find_elements(By.LINK_TEXT,'Xem')) > 0:
		driver.find_element(By.LINK_TEXT,'Xem').click()
		time.sleep(1.75)
		raw = driver.find_elements(By.CLASS_NAME,'jrPage')[1]
		rows = raw.find_elements(By.TAG_NAME, "tr") # get all of the rows in the table
		for row in rows:
			col = row.find_elements(By.TAG_NAME, "td")
			if (len(col) == 4) and (col[0].text != "Tổng cộng") and (col[0].text != "Loại hình giao dịch"):
				table_row = []
				for data in col:
					table_row.append(data.text)
				result_table.append(table_row)
		r = 0
		position = {}
		save_position = []
		while True:
			if (result_table[r][0].find('Mua') != -1) or (result_table[r][0].find('Bán') != -1):
				temp = result_table[r][0]
				save_position.append(r)
			position.update({r:temp})
			r+=1
			if r >= len(result_table):
				break
		table_data = []
		for r in range(len(result_table)):
			row_data = []
			if r not in save_position:
				row_data.append(date.date())
				row_data.append(position[r])
				row_data.append(int(result_table[r][0].split()[-1]))
				row_data.append(result_table[r][1])
				row_data.append(float(result_table[r][2].replace('.','').replace(',','.')))
				row_data.append(float(result_table[r][3].replace('.','').replace(',','.')))
				table_data.append(row_data)

		return table_data
	driver.quit()

wb = openpyxl.load_workbook(home+'vn_omo.xlsx')
ws = wb.active

latest = ws.cell(row=len(ws['A']),column=1).value+datetime.timedelta(1)

if latest>today:
	quit()

r=len(ws['A'])+1

while True:
	print(latest)
	if getmmdata(latest):
		a = getmmdata(latest)
		for i in range(len(a)):
			c = 1
			for j in range(len(a[i])):
				ws.cell(row = r+i, column=c+j, value=a[i][j])
		r = r+len(a)

	latest = latest + datetime.timedelta(1)

	wb.save(home+'vn_omo.xlsx')
	if latest > today:
		break