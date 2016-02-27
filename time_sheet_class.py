from datetime import datetime

Class TimeWorked:
	
	# initializer
	def __init__(inTime, outTime, totalTime, fileName, currentCell):

		self.inTime = inTime
		self.outTime = outTime
		self.totalTime = 0.0 
		self.fileName = fileName 
		self.currentCell = currentCell

	# calculate hours worked
	def calculate_hours_worked(self):

		#time format
		FMT = '%H:%M:%S'
		s1 = self.inTime
		s2 = self.outTime
		tDelta = datetime.strptime(s2, FMT) - datetime.strptime(s1, FMT)
	
	
