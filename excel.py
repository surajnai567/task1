import pyexcel


# this function reads excel sheet row by row with pyexcel module
def get_excel_sheet(filename, sheet_number=0):
	sheet = pyexcel.get_sheet(file_name=filename, sheet_name=0)
	sheet.name_columns_by_row(0)
	for i, row in enumerate(sheet):
		if row[3] == 1:
			yield i, row
		else:
			continue


def update_to_excel(filename, row, data, sheet_number=0):
	sheet = pyexcel.get_sheet(file_name=filename, sheet_name=sheet_number)
	sheet.row[row+1] = data
	sheet.save_as(filename)

