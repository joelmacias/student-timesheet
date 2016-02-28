from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
import sys
import new_workbook
import open_workbook
from time_sheet_class import TimeSheet
# prompt for new sheet or load sheet
print "Open existing workbook type"
userWorkbookSelection = raw_input()


''' User has selected to create new workbook. Prompt for name of file to be created.
	Instance of class TimesSheet is created with arguments fileName, and a workbook 
	returned by the function create_new_workbook.
	
	create_new_workbook will format the sheet to mimic the standard student worker timesheet 
'''
if userWorkbookSelection == "n":
	
	# get file name
	print "Enter a filename for new workbook: "
	fileName = raw_input()
	
	# instance of class TimeSheet
	currentTimeSheet = TimeSheet(fileName, new_workbook.create_new_workbook(fileName))



if userWorkbookSelection == "o": 

	print "Enter the name of the file to be opened: "
	fileName = raw_input()
	# wb is the current workbook
	currentTimeSheet = TimeSheet(fileName,open_workbook.open_existing_workbook(fileName))


# find row where user input will start
currentTimeSheet.find_empty_cell()

# 22 last cell that can be written to A24
if currentTimeSheet.cellIndex == -1:
	sys.exit("Error: Timesheet is full. Create new Timesheet")
else: 
	print "Add a shift entry? (y/n): "
	addNewEntry = raw_input()

	if addNewEntry == 'n':	
		currentTimeSheet.calculate_grand_total()
		currentTimeSheet.save_time_sheet()
		sys.exit("Saving timesheet, and exiting program.")


	while (addNewEntry == 'y'):
		shiftInfoList = open_workbook.get_shift_info()
		print "The following will be input:\n Date: {0}\nIn time: {1}\nOut time: {2}\nTotal time worked: {3}".format(shiftInfoList[0], shiftInfoList[2], shiftInfoList[4], shiftInfoList[5])
		
		print "Discard shift entry, and re-enter shift entry? (y/n): "
		discardInput = raw_input()
		
		if discardInput == 'n':
			currentTimeSheet.insert_entry(shiftInfoList)
			currentTimeSheet.find_empty_cell()
			print "Shift entry has been added." 
		
		if discardInput == 'n' and currentTimeSheet.cellIndex != -1:
			print "Add another entry? (y/n): "
			addNewEntry = raw_input()

		if currentTimeSheet.cellIndex == -1 and addNewEntry == 'y': 
			currentTimeSheet.calculate_grand_total()
			currentTimeSheet.save_time_sheet()
			sys.exit("The current timesheet is full, saving timesheet, and exiting program.") 
				
		if discardInput == 'n' and addNewEntry == 'n':
			currentTimeSheet.calculate_grand_total()
			currentTimeSheet.save_time_sheet()
			sys.exit("Saving timesheet, and exiting program.")


