import requests


def get_temperature_in_kelvin(city_name, url, apikey):
	res = requests.get(url+city_name+apikey)
	return res.json()['main']['temp']


# api give temp in kelvin fun to convert into c or f
def convert_temperature(temp_in_kelvin, convert_to):
	celsius = temp_in_kelvin - 273.15
	if convert_to.lower() == 'c':
		return celsius
	if convert_to.lower() == 'f':
		return celsius*9/5 + 32


# update the excel sheet in real time with xlwings
def update(sheet, index, value):
	sheet.range(index).value = value


