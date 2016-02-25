from openpyxl import Workbook
from openpyxl import load_workbook
	
def open_existing_workbook():

	fileExtension = ".xlsx"
	# get name of file to be opened 
	print "Enter the name of the file you want to open"
	fileToOpen = raw_input() 
	fileToOpen += fileExtension
	
	# load workbook from given filename
	return load_workbook(filename = fileToOpen)

