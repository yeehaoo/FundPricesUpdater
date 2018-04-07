import bs4
import openpyxl
import xlrd
import xlwt
from xlutils.copy import copy

#change file names as desired
htmlFileName = 'webpage.html'
workbookName = 'Quickswitch.xls'

#parsing the html of the webpage
file = open(htmlFileName)
soup = bs4.BeautifulSoup(file, "html.parser")

#initialising values and valuesOffer arrays, and getting all table cell values
elem = soup.select('#Form1 td')
values = ['']
valuesOffer = ['']

#getting the wanted values and storing it in array
for i in range(20):
        value = elem[5*i + 7].getText()
	values.append(value)
	valuesOffer.append(str(float(value) * 1.035))

#remove the initialisation value from arrays
values.remove('')
valuesOffer.remove('')

#open workbook and read max row number
rb = xlrd.open_workbook(workbookName)
maxRowNo = int(rb.sheet_by_index(0).nrows) + 1

#make a writable copy of workbook
wb = copy(rb)

#select the first sheet to write on
w_sheet = wb.get_sheet(0)

#write values to sheet
for i in range(20):
	w_sheet.write(maxRowNo,i,values[i])

#save the workbook
wb.save('Updated.xls')
