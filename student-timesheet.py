from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import new_workbook
import open_workbook

# prompt for new sheet or load sheet
print "Usage - Enter 'n' to create a new excel workbook, enter 'o' to open an existing excel workbook"
userWorkbookSelection = raw_input()


# user has selected to create a new workbook
if userWorkbookSelection == "n":
	
	# wb is the current workbook
	wb = new_workbook.create_new_workbook() 
	# get worksheet
	ws = wb.active

	 
if userWorkbookSelection == "o": 

	# wb is the current workbook
	wb = open_workbook.open_existing_workbook()
	# get worksheet	
	ws = wb.active

'''
	check = 'A7'	
	emptyCheck = ws[''].value
	if emptyCheck == None: 
		print 'empty'
'''

# find row where user input will start
indexOfCellToBeWritten = new_workbook.find_empty_cell(ws)

# 22 last cell that can be written to A24
if indexOfCellToBeWritten == -1:
	print "Sheet is full" 

if indexOfCellToBeWritten  != -1:
	print "Add an entry? (y/n): "
	addNewEntery = raw_input() 
	while (addNewEntery == 'y'):
		print "Enter date <dd/mm/yy> : "
		currentDate = raw_input()
		print "Enter in time - <HH:MM> : "
		inTime = raw_input()
		print "Enter out time - <HH:MM> : "
		outTime = raw_input() 
		FMT = '%H:%M'
		timeDelta = datetime.strptime(outTime, FMT) - datetime.strptime(inTime, FMT) 
		print "The following will be input:\n Date: {0}\nIn time: {1}\nOut time: {2}\nTotal time worked: {3}".format(currentDate, inTime, outTime, timeDelta)
		print "Edit input? (y/n): "
		editInput = raw_input()
		if editInput == 'n':
			x = 7
		# go to insert function
		print "Add another entry? (y/n): "
		addNewEntry = raw_input()

'''
#print (ws.columns)[0]
for i in (ws.columns)[0]:
	if i.value == None:
		print i
		#print " {0}".format(x)
		#x = x + 1; 
'''

