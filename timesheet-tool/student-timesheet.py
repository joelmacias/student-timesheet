from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime
from argparse import ArgumentParser
from time_sheet_class import TimeSheet
import argparse
import new_workbook
import open_workbook
import sys


def main(userWorkbookSelection, fileName):
		
		# New timesheet will be created.	
		if userWorkbookSelection == "new":
			
			# CurrentTimeSheet is an instance of class TimeSheet. 
			# Create_new_workbook will format the timesheet to mimic the standard student worker timesheet.
			currentTimeSheet = TimeSheet(fileName, new_workbook.create_new_workbook(fileName))
		
		# Existing timesheet will be opened.
		if userWorkbookSelection == "open": 
			currentTimeSheet = TimeSheet(fileName,open_workbook.open_existing_workbook(fileName))


		# Find row in timesheet where shift entry will be inserted.
		# find_empty_cell() will set class variable cellIndex to index where. 
		# Shift entry will be inserted, or -1 if timesheet is full.
		currentTimeSheet.find_empty_cell()

		# If the opened timesheet is full exit program.
		# If not, prompt for shift entry.
		if currentTimeSheet.cellIndex == -1:
			sys.exit("Error: Timesheet is full. Create new Timesheet")
		else: 
			print "Add a shift entry? (y/n): "
			addNewEntry = raw_input()

			# Calculate total, save timesheet, exit program.
			if addNewEntry == 'n':	
				currentTimeSheet.calculate_grand_total()
				currentTimeSheet.save_time_sheet()
				sys.exit("Saving timesheet, and exiting program.")

			# Keep looping until user no longer wants to add shift entries. 
			while (addNewEntry == 'y'):
				
				# shiftInfoList will store the shift entry: date, int time label, in time, out time labelout time, and total hours worked. 
				shiftInfoList = open_workbook.get_shift_info()
				print "\nYou entered the following shift:\nDate: {0}\nIn time: {1}\nOut time: {2}\nTotal time worked: {3}".format(shiftInfoList[0], shiftInfoList[2], shiftInfoList[4], shiftInfoList[5])
				
				print "\nSave shift entry, or discard shift entry? (s/d): "
				discardInput = raw_input()
				
				# Shift entry will be inserted into timesheet, next cell to be written to will be located
				# and timesheet will be saved.
				if discardInput == 's':
					currentTimeSheet.insert_entry(shiftInfoList)
					currentTimeSheet.find_empty_cell()
					currentTimeSheet.save_time_sheet()
					print "Shift entry has been added." 
				
				# Prompt user for another entry.
				if (discardInput == 'd' or discardInput == 's') and currentTimeSheet.cellIndex != -1:
					if discardInput =='d':
						print "Entry has been discared"
					print "\nAdd another entry? (y/n): "
					addNewEntry = raw_input()

				# If timesheet is full and user wants to add another shift, calculate grand total for hours worked
				# save timesheet, and exit program.
				if currentTimeSheet.cellIndex == -1 and addNewEntry == 'y': 
					currentTimeSheet.calculate_grand_total()
					currentTimeSheet.save_time_sheet()
					sys.exit("The current timesheet is full, saving timesheet, and exiting program.") 
				
				# User is done entering shift data, calcualte grand total of hours worked, save timesheet,exit program
				if discardInput == 'd' or discardInput == 's' and addNewEntry == 'n':
					currentTimeSheet.calculate_grand_total()
					currentTimeSheet.save_time_sheet()
					sys.exit("Saving timesheet, and exiting program.")

if __name__ == "__main__":
	parser = argparse.ArgumentParser()
	parser.add_argument("timesheet", choices=['open', 'new'], help = "o = open existing timesheet\nn = create new timesheet")
	parser.add_argument("filename", help = "Name of file to open, or name of file to create")
	args = parser.parse_args()
	main(args.timesheet, args.filename) 
