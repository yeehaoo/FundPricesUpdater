import bs4
import openpyxl
import xlrd
import xlwt
import datetime
from xlutils.copy import copy

#change file names as desired
htmlFileName = 'web2.html'
workbookName = 'fundprices14.xls'

#parsing the html of the webpage
file = open(htmlFileName)
soup = bs4.BeautifulSoup(file.read(), "html.parser")

#initialising values and valuesOffer arrays, and getting all table cell values
elem = soup.select('#Form1 td')
values = [0.01]
valuesOffer = [0.01]

#getting the wanted values and storing it in array
for i in range(21):
	value = float(elem[5*i + 7].getText())
	values.append(value)
	valuesOffer.append(float(value) * 1.035)

#remove the initialisation value from arrays
values.remove(0.01)
valuesOffer.remove(0.01)

#sort arrays
#correct order: 14,7,8,5,1,0,16,15,3,6,2,4,9,10,11,12,19,17,13,18
values = [values[14],values[7],values[8],values[5],values[1],values[0],values[16],values[15],values[3],values[6],values[2],values[4],values[9],values[10],values[11],values[12],values[19],values[17],values[13],values[18]]
for i in range(20):
	temp = float(values[i]) * 1.035
	valuesOffer[i] = temp

#open workbook and read max row number
rb = xlrd.open_workbook(workbookName)
maxRowNo = int(rb.sheet_by_index(0).nrows)


#make a writable copy of workbook
wb = copy(rb)

#select the first sheet to write on
w_sheet = wb.get_sheet(0)

#getting today's date and writing it to the first row
todaysDate = datetime.date.today()
todaysDate = '/'.join([str(todaysDate.day),str(todaysDate.month)])
w_sheet.write(maxRowNo,0,todaysDate)

#writing values to sheet
#rows for values array: 5,7,9,11,13,15,17,21,23,25,27,29,34,38,40,42,44,46,48,50
rowsToWrite = [5,7,9,11,13,15,17,21,23,25,27,29,34,38,40,42,44,46,48,50]
for i in range(20):
	w_sheet.write(maxRowNo,rowsToWrite[i],values[i])
	w_sheet.write(maxRowNo,rowsToWrite[i] + 1,valuesOffer[i])

"""
#write values to sheet
for i in range(20):
	w_sheet.write(maxRowNo,i,values[i])
"""

#save the workbook
wb.save('Updated.xls')
