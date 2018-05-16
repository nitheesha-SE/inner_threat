	
import os
import os.path
import xlrd
import xlsxwriter
for i in range(1,41):

	num = '%0.2d' % i
	file_name = "c://mouse_4feature//sub_together//" + num
	merged_file_name =file_name + ".xlsx"

	if merged_file_name == "c://sample//R40031.xlsx": 
		file_name = "c://sample//" + 03
		merged_file_name =file_name + ".xlsx"	
		dest_book = xlsxopen.Workbook(merged_file_name)
		dest_book = openpyxl.load_workbook(merged_file_name)
		dest_sheet_1 = dest_book.worksheets[0]
		dest_row = dest_sheet_1.max_row
	elif merged_file_name == "c://sample//R40034.xlsx": 
		file_name = "c://sample//" + '08'
		merged_file_name =file_name + ".xlsx"	
		dest_book = xlsxopen.Workbook(merged_file_name)
		dest_book = openpyxl.load_workbook(merged_file_name)
		dest_sheet_1 = dest_book.worksheets[0]
		dest_row = dest_sheet_1.max_row
	else:
		dest_book = xlsxwriter.Workbook(merged_file_name)
		dest_sheet_1 = dest_book.add_worksheet()
		dest_row = 1
	temp = 0
	path = "C://mouse_4feature//all_together//"
	print ("entering")
	for root,dirs,files in os.walk(path):
		print ("entering")
		files = [ _ for _ in files if _.endswith( num +'.xlsx') ]
		for xlsfile in files:
			print ("File in mentioned folder is: " + xlsfile)
			temp_book = xlrd.open_workbook(os.path.join(root,xlsfile))
			temp_sheet = temp_book.sheet_by_index(0)
			ncols = temp_sheet.ncols
			nrows = temp_sheet.nrows
			if temp == 0:
				for col_index in range(temp_sheet.ncols):
					str = temp_sheet.cell_value(0, col_index)
					dest_sheet_1.write(0, col_index, str)
					temp = temp + 1
			for row_index in range(1, nrows):
				for col_index in range(ncols):
					str = temp_sheet.cell_value(row_index, col_index)
					dest_sheet_1.write(dest_row, col_index, str)
				dest_row = dest_row + 1
	dest_book.close()
	book = xlrd.open_workbook(merged_file_name)
	sheet = book.sheet_by_index(0)
	print "number of rows in destination file are: ", sheet.nrows
	print "number of columns in destination file are: ", sheet.ncols