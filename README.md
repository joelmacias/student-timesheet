#Student Worker Timesheet
Command line tool (CLT) to generate my student worker timsheet. The main goal of this CLT was to start learning Python, and to deter me from filling out all my paper timsheets the morning they are due. My intention is to fire up the tool at the start, and end of my shift, logging times accordingly.

I'm certainly aware that I could have simply created this timesheet using Excel, but that would have been no fun.
#Usage
The CLT will allow you to open an existing timesheet, or to create a new timesheet. 
The format of the timesheet created will be the standard timesheet provided by my department. The tool will prompt for worker information, then proceed to prompt for shift entries.  

I wrote this tool specifically for the timesheet provided by my department, so it wont work properly with any other timesheet. :)

To run the tool:

	python student-timesheet.py <new|open> <filename>
	
First argument: < new | open >  
'new' will create a new timesheet, and format it according to the department timesheet.
'open' will open an existing timesheet.

Second argument: < filename > 
A new timesheet will be created with the provided name if the first argument was new.
An existing timesheet will be opened if the first argument was open 

##Example timesheet 
![Example timesheet](https://raw.githubusercontent.com/joelmacias/student-timesheet/master/sample_timesheet.jpg)
