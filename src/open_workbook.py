#open_workbook.py
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from datetime import timedelta
import os
import sys

def open_existing_workbook(fileName):
	""" Returns a workbook from the filename specified by the user """
	raw = fileName;

#	fileExtension = ".xlsx"
#   fileName += fileExtension

	# current path
	pwdPath = os.getcwd()
	pwdPath = pwdPath + "/" + fileName
	if(os.path.isfile(pwdPath)):
		print ("File '{0}' has been opened".format(fileName)) 	
		# Load workbook from given filename.
		return load_workbook(filename = fileName)
	else:
		print ("File '{0}' does not exist.\nExiting program.".format(raw)) 
		sys.exit() 

def standard_time_to_military_time(rawTimeInput):
	# tuple with in time and am/pm [0] = intime, [1] = ' ', [2] = am/pm
	inTimeTuple = rawTimeInput.partition(' ') 
	
	# inTimeTuple[0] = hour, inTimeTuple[1] = ':', inTimeTuple[2] = min	
	hourAndMinTuple = inTimeTuple[0].partition(':')
	

	if inTimeTuple[2] == 'pm':
		# if less than 12 convert to military time by adding 12 
		if int(hourAndMinTuple[0]) < 12:
			inTimeMilitaryHour = int(hourAndMinTuple[0]) + 12 
		else:
			inTimeMilitaryHour = 12
	
		inTimeFormatted = str(inTimeMilitaryHour) + ':' + hourAndMinTuple[2]

	if inTimeTuple[2] == 'am':
		# if current hour is 12am change it to military time 12am which is 00	
		if int(hourAndMinTuple[0]) == 12:
			inTimeFormatted = '00' + ':' + hourAndMinTuple[2]		
		else:
			inTimeFormatted = hourAndMinTuple[0] + ':' + hourAndMinTuple[2]

	return inTimeFormatted

def get_shift_info():
	""" Returns a list with the shift data entered by the user, and labels to add to timesheet.
		The list will have the following data: shift date, in time label, in time, out time label, out time, and total hours worked for the shfit """
	
	# Create an empty list.
	shiftInfoList = []

	# Prompt for shif date, appened date to list, add in time label to list.
	print ("\nEnter date <dd/mm/yy>: ")
	currentDate = input()
	shiftInfoList.append(currentDate)
	shiftInfoList.append("IN TIME:")
	
	# Prompt user for in time, appened in time to list, add out time label.
	print ("Enter in time - <HH:MM am/pm>: ")
	
	inTimeRawInput = input()
	shiftInfoList.append(inTimeRawInput)
	shiftInfoList.append("OUT TIME:")

	inTime = standard_time_to_military_time(inTimeRawInput)


	# Prompt user for out time, append out time to list.	
	print ("Enter out time - <HH:MM am/pm>: ")
	outTimeRawInput = input()
	shiftInfoList.append(outTimeRawInput)

	outTime = standard_time_to_military_time(outTimeRawInput) 
	
	# Calculate total time worked in shift.  
	FMT = '%H:%M'
	timeWorked = datetime.strptime(outTime, FMT) - datetime.strptime(inTime, FMT)

	# Out time will always be assumed to be later than in time, so remove negative days
	if timeWorked.days < 0:
		timeWorked = timedelta(days = 0, seconds = timeWorked.seconds, microseconds = timeWorked.microseconds)	

	# Appened total time worked for shift to list as a string 
	shiftInfoList.append(str(timeWorked))
	
	return shiftInfoList


