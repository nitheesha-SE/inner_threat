import os,sys,csv
import xlsxwriter

def processFile(path):
	folder = path
	#read ech file from directory
	for filename in os.listdir(folder):
		#save to new directory
		address = "E:\\"
		file_name = address + filename + '.xlsx'
		#create a new worksheet in write mode
		workbook = xlsxwriter.Workbook(file_name)
		worksheet = workbook.add_worksheet()
		print("in"+filename)	
		#loop to each file int he folder
		infilename = os.path.join(folder,filename)
		#open the csv file
		reader = csv.reader(open(infilename , "rb"), delimiter = '|')
		for r, row in enumerate(reader):
			for c, col in enumerate(row):
				#write each cell value to spreadsheet
				worksheet.write(r, c, col)
	
		workbook.close()
#get the file name from command promt
processFile(sys.argv[1])