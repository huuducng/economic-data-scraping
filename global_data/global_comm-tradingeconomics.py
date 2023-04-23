from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

from bs4 import BeautifulSoup
import time, datetime, openpyxl, requests, pathlib
import pandas as pd

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

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://tradingeconomics.com/commodities")

page_source = driver.page_source
soup = BeautifulSoup(page_source, 'html.parser')

tables = soup.find_all('table')

driver.quit()

wb = openpyxl.Workbook()
ws = wb.active

r = 1

for table in tables:
    table = tableDataText(table)
    for row in table:
        c = 0
        for data in row:
            c+=1
            ws.cell(row=r, column=c, value=data)
        r+=1
    r+=1

wb.save('global_comm.xlsx')