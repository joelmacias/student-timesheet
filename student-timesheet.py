from openpyxl import Workbook
from openpyxl import load_workbook

import new_workbook
import open_workbook

# prompt for new sheet or load sheet
print "To create a new workbook enter 'new'\nTo open an existing timesheet enter 'open'"
userWorkbookSelection = raw_input()


# user has selected to create a new workbook
if userWorkbookSelection == "new":
	
	# wb is the current workbook
	wb = new_workbook.create_new_workbook() 
	# get worksheet
	ws = wb.active
	print ws.cell('A3').value 
	 
if userWorkbookSelection == "open": 

	# wb is the current workbook
	wb = open_workbook.open_existing_workbook()
	# get worksheet	
	ws = wb.active


