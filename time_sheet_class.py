#time_sheet_class.py
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import time 
class TimeSheet:
	""" A time sheet structured to mimic the standard student worker time sheet at UNLV 


	Attributes:
		cellIndex 			Next cell where shift data should be written to
		totalHoursWorked	Total hours worked when calculating time worked for all shifts
		totalMinWorked		Total minutes worked when calculating time worked for all shifts
	"""

	cellIndex = 5
	totalHoursWorked = 0
	totalMinWorked = 0

	# initializer
	def __init__(self, fileName, wb):
		self.fileName = fileName
		self.wb = wb
		self.ws = wb.active
		(self.ws).print_grid = True

	def find_empty_cell(self):
		""" Finds next empty cell where shift data can be written to. If no cells remanin then cellIndex is set to -1 """
		for i, x in enumerate(((self.ws).columns)[0]):
			if i >= 5:
				if x.value == None:
					self.cellIndex = i
					return 
		self.cellIndex = -1

	def insert_entry(self,entryList):
		""" Uses the shift data in the list entryList to populate the current row determined by cell inded """
		for i, x in enumerate(((self.ws).rows)[self.cellIndex]):
			x.value = entryList[i]
			
	def save_time_sheet(self):
		""" Saves the time sheet with the file name provided by the user """
		fileExtension = ".xlsx"
		fileToSave = self.fileName
		fileToSave += fileExtension 
		self.wb.save(fileToSave)
	
	def calculate_grand_total(self):
		""" Adds all the hours worked to determine total time worked for the time sheet, and populates cell with the grand total """
		grandTotal = '00:00:00'
		FMT = '%H:%M:%S'	
		y = datetime.datetime.strptime(grandTotal,FMT).time()

		for x in (((self.ws).columns)[5])[5:22]:
			if x.value != None:
				self.totalHoursWorked += (datetime.datetime.strptime(x.value,FMT).time()).hour
				self.totalMinWorked +=  (datetime.datetime.strptime(x.value,FMT).time()).minute
 
		while(self.totalMinWorked >= 60):
				self.totalMinWorked =  self.totalMinWorked - 60
				self.totalHoursWorked += 1

		grandTotal = datetime.time(self.totalHoursWorked, self.totalMinWorked, 00)
 		(self.ws).cell('F24').value = grandTotal


		
