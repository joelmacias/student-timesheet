from datetime import datetime

Class TimeWorked:
	
	# initializer
	def __init__(self, date, inTime, outTime, totalTime):

		self.name = name
		self.date = date
		self.inTime = inTime
		self.outTime = outTime
		self.totalTime = 0.0 

	# calculate hours worked
	def calculate_hours_worked(self):

		#time format
		FMT = '%H:%M:%S'
		s1 = self.inTime
		s2 = self.outTime
		tDelta = datetime.strptime(s2, FMT) - datetime.strptime(s1, FMT)
	
	# return class variables
	def return_in_time(self): return self.inTime
	def return_out_time(self): return self.outTime
	def return_total_time(self): return  self.totalTime


