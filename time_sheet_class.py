#time_sheet_class.py
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import time 
class TimeSheet:
	
	cellIndex = 5
	totalHoursWorked = 0
	totalMinWorked = 0

	# initializer
	def __init__(self, fileName, wb):
		self.fileName = fileName
		self.wb = wb
		self.ws = wb.active

	# calculate hours worked
	def calculate_hours_worked(self):
		#time format
		FMT = '%H:%M:%S'
		s1 = self.inTime
		s2 = self.outTime
		tDelta = datetime.datetime.strptime(s2, FMT) - datetime.datetime.strptime(s1, FMT)
	
	def find_empty_cell(self):
		for i, x in enumerate(((self.ws).columns)[0]):
			if i >= 5:
				if x.value == None:
					self.cellIndex = i
					return 
		self.cellIndex = -1

	def insert_entry(self,entryList):
		for i, x in enumerate(((self.ws).rows)[self.cellIndex]):
			x.value = entryList[i]
			#print x
	def save_time_sheet(self):
		fileExtension = ".xlsx"
		fileToSave = self.fileName
		fileToSave += fileExtension 
		self.wb.save(fileToSave)
	
	def calculate_grand_total(self):
		grandTotal = '00:00:00'
		FMT = '%H:%M:%S'
		check = False
		z = 0.0 
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
		print grandTotal

		
