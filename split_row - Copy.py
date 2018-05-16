import xlrd
import openpyxl
import xlsxwriter
import os,sys
from xlutils.copy import copy
file_name = "E:\\sample.xlsx"
wb = xlrd.open_workbook(file_name)
sheet = wb.sheet_by_index(0)
nrows = sheet.nrows
ncol = sheet.ncols
c=0
i = 1
max_row = nrows/10
address = "E:\\"
#max sheets with 10 rows each
while (i<=max_row):
	#create a new sheet for each set of values
    filename = address + str(i) + '.xlsx'
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    p = 1
	#num of rows in a sheet
    while(p <=10) and c < nrows:
        col= 0
		#loop for each row
        while (col <ncol):
			#get the cell values
			data = sheet.cell_value(c,col)
			data1 = sheet.cell_value(0,col)
			worksheet.write(p,col,data)
			worksheet.write(0,col,data1)
			col += 1
        print "one row compleeeted"
        p += 1
        c += 1
    print "one row set completed"
    i += 1
    workbook.close()
