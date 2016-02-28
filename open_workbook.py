#open_workbook.py
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from datetime import timedelta	
def open_existing_workbook(fileName):
	fileExtension = ".xlsx"
	fileName += fileExtension

	print "File '{0}' has been opened".format(fileName) 	
	# load workbook from given filename
	return load_workbook(filename = fileName)

def get_shift_info():
	# create an empty list
	shiftInfoList = []

	print "Enter date <dd/mm/yy> : "
	currentDate = raw_input()
	shiftInfoList.append(currentDate)
	shiftInfoList.append("IN TIME:")
	
	print "Use 24-hour time format" 
	print "Enter in time - <HH:MM> : "
	inTime = raw_input()
	shiftInfoList.append(inTime)
	shiftInfoList.append("OUT TIME:")

	print "Enter out time - <HH:MM> : "
	outTime = raw_input()
	shiftInfoList.append(outTime)
	
	FMT = '%H:%M'
	timeWorked = datetime.strptime(outTime, FMT) - datetime.strptime(inTime, FMT)
	#print ((datetime.strptime(outTime, FMT) - datetime.strptime(inTime, FMT)).time()).hour

	if timeWorked.days < 0:
		timeWorked = timedelta(days = 0, seconds = timeWorked.seconds, microseconds = timeWorked.microseconds)	

	#str(timeWorked)
	shiftInfoList.append(str(timeWorked))
	
	return shiftInfoList


