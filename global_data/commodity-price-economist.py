import requests, datetime
from bs4 import BeautifulSoup
from PIL import Image

delta = input('Number of previous weeks (Current week: 0): ')

today = datetime.date.today()
start_date = today + datetime.timedelta( (5-today.weekday()) % 7 ) + datetime.timedelta(-7*int(delta))

for i in range(7):
	date = start_date + datetime.timedelta(-i)
	targeturl = date.strftime('%Y/%m/%d')
	site = 'https://www.economist.com/economic-and-financial-indicators/'+targeturl+'/economic-data-commodities-and-markets'
	response = requests.get(site)
	if response:
		break

# if targeturl == today.strftime('%Y/%m/%d'):
# 	print('Currently, data is not available for selected week!\nPlease increase number of previous weeks')
# 	quit()

print(targeturl)

targetname = start_date.strftime('%Y%m%d')+'_INT401.png'
savename = start_date.strftime('%Y-%m-%d-Commoditiy-price Index - Economist.png')


soup = BeautifulSoup(response.text, 'html.parser')
img_tags = soup.find_all('img')

targetfile=''
urls = [img['src'] for img in img_tags]
for item in urls:
	if targetname.lower() in item.lower():
		targetfile = item

if targetfile =='':
	print('Please check mannually!')
	quit()

file = requests.get(targetfile)
open(savename,'wb').write(file.content)
Image.open(savename).show()