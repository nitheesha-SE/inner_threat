import pandas as pd
from pandas import ExcelWriter
import os
#source file address
address = "C:\\key_feature\\100division\\alltogether\\subject_together\\"
#destination file address
result = "C:\\key_feature\\100division\\alltogether\\subject_together\\key\\"
directory = os.listdir(address)
#loop for each file 
for filename in directory:
	output = filename
	filename = address+filename
	print (filename)
	print (output)
	#get the file name
	wt = result+output+".xlsx"
	#read the specified columns
	f = pd.read_excel(filename , index_col=None, na_values=['NA'], parse_cols ="u:bf")
	#write to new file
	writer=pd.ExcelWriter(wt, engine='xlsxwriter')
	f.to_excel(writer,sheet_name ='sheet1')
	writer.save()

