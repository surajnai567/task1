import excel
import util
import time
import xlwings

# global variables
url = "http://api.openweathermap.org/data/2.5/weather?q="
apikey = "&appid=" # put your api key after = sign
update = True
update_freq = 3  # set update frequency
filename = "data.xlsx" # provide the file location here

wb = xlwings.Book(filename)
sheet1 = wb.sheets['sheet1']


while update:
	data = []
	# read sheet row-wise
	for i, row in excel.get_excel_sheet(filename):
		temp = util.get_temperature_in_kelvin(row[0], url, apikey)
		temp = util.convert_temperature(temp, row[2])
		row[1] = temp
		data.append([i+2, row])
	# iterate over data and fill the excel sheet in real time
	for dat in data:
		# print("B{}".format(dat[0]), dat[1][1])
		util.update(sheet1, "B{}".format(dat[0]), dat[1][1])

	time.sleep(update_freq)


