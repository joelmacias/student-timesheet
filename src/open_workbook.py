#open_workbook.py
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from datetime import timedelta

	
def open_existing_workbook(fileName):
	""" Returns a workbook from the filename specified by the user """

	fileExtension = ".xlsx"
	fileName += fileExtension

	print "File '{0}' has been opened".format(fileName) 	
	# Load workbook from given filename.
	return load_workbook(filename = fileName)

def get_shift_info():
	""" Returns a list with the shift data entered by the user, and labels to add to timesheet.
		The list will have the following data: shift date, in time label, in time, out time label, out time, and total hours worked for the shfit """
	
	# Create an empty list.
	shiftInfoList = []

	# Prompt for shif date, appened date to list, add in time label to list.
	print "\nEnter date <dd/mm/yy>: "
	currentDate = raw_input()
	shiftInfoList.append(currentDate)
	shiftInfoList.append("IN TIME:")
	
	# Prompt user for in time, appened in time to list, add out time label.
	print "\nUse 24-hour time format" 
	print "Enter in time - <HH:MM>: "
	inTime = raw_input()
	shiftInfoList.append(inTime)
	shiftInfoList.append("OUT TIME:")

	# Prompt user for out time, append out time to list.
	print "\nUse 24-hour time format"
	print "Enter out time - <HH:MM>: "
	outTime = raw_input()
	shiftInfoList.append(outTime)
	
	# Calculate total time worked in shift.  
	FMT = '%H:%M'
	timeWorked = datetime.strptime(outTime, FMT) - datetime.strptime(inTime, FMT)

	# Out time will always be assumed to be later than in time, so remove negative days
	if timeWorked.days < 0:
		timeWorked = timedelta(days = 0, seconds = timeWorked.seconds, microseconds = timeWorked.microseconds)	

	# Appened total time worked for shift to list as a string 
	shiftInfoList.append(str(timeWorked))
	
	return shiftInfoList


