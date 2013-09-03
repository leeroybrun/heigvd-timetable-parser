import sys
import datetime
import math

import xlrd
from icalendar import Calendar, Event, LocalTimezone

# LINKS :
#  - Save to .ics : http://icalendar.readthedocs.org/en/latest/
#       http://icalendar.readthedocs.org/en/latest/examples.html
#       http://icalendar.readthedocs.org/en/latest/examples.html#more-documentation
#       
#  - Read xls : 
#       https://github.com/python-excel/xlrd
#       http://scienceoss.com/read-excel-files-from-python/
#       https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966
#       http://www.youlikeprogramming.com/2012/03/examples-reading-excel-xls-documents-using-pythons-xlrd/
#       
#       
#       
#       

# Get start date of week, and then increment when changing day
#Date = datetime.datetime.strptime(StartDate, "%m/%d/%y")
#EndDate = Date + datetime.timedelta(days=10)

###########################################################
# Position of data inside the XLS sheet / global vars
###########################################################
START_DATE_COL = 0 # Column num of the week start date
START_COL = 2      # Start column of data
HOURS_ROW = 2      # Row containing hours
START_ROW = 3      # Start row of data
WEEK_ROWS = 3      # Number of rows per week
NB_DAYS = 6


###########################################################
# Timetable class
###########################################################

class Timetable:
	def __init__(self, file):
		self.file = file

		# TODO: instanciate an ICS object

	# Function used to parse a timetable xls file
	def parse(self):
		self.wb = xlrd.open_workbook(self.file, formatting_info=True)
		self.sh = self.wb.sheet_by_index(0)

		self.parseHours(HOURS_ROW)

		for rowNum in range(START_ROW, self.sh.nrows, WEEK_ROWS):
			self.parseRow(rowNum)

	# Parse rows of XLS file (which in fact is a week)
	def parseRow(self, rowNum):
		print 'parseRow ', rowNum

		sh = self.sh
		wb = self.wb

		# If start date of week is not a date, return. Can't proceed it. (usually not a valid week row)
		if sh.cell_type(rowNum, START_DATE_COL) != 3:
			return;

		# Get start date of the week
		startDateData = xlrd.xldate_as_tuple(sh.cell_value(rowNum, START_DATE_COL), wb.datemode)
		weekStartDate = datetime.datetime(startDateData[0], startDateData[1], startDateData[2], 0, 0, 0)

		nbColsPerDay = len(self.hours)+1
		dayNum = 0

		# Loop over each days
		for startDayCol in range(START_COL, sh.row_len(rowNum), nbColsPerDay):
			if dayNum >= NB_DAYS:
				break;

			# Check if first day cell has color of "vacation" day
			if self.isVacation():
				# TODO: add day to ics with summary "Congé"
				continue;

			# Ignore first col of day (contains the day.month)
			startDayCol += 1

			print 'endDayCol ', startDayCol+len(self.hours)

			# TODO: parse courses and add them to ics
			# Loop over each courses
			#for courseCol in range(startDayCol, startDayCol+len(self.hours)):
				#print 'courseCol ', courseCol

			dayNum += 1

	def isVacation(self, row, col):
		print 'startDayCol ',startDayCol, ' - rowNum ',rowNum
		xfIndex = sh.cell_xf_index(startDayCol, rowNum)
		print 'xfindex ', xfIndex
		if xfIndex:
			cellXf = wb.xf_list[sh.cell_xf_index(startDayCol, rowNum)]

			print cellXf

	# Get start hours of courses
	def parseHours(self, rowNum):
		sh = self.sh

		# 0 = not run, 1 = first empty cell met, 2 = second empty step met (STOP)
		emptyCells = 0
		colNum = START_COL

		self.hours = []

		while emptyCells != 2:
			cellType = sh.cell_type(rowNum, colNum)
			if cellType == 0 or cellType == 6:
				emptyCells += 1
			else:
				hourFloat = sh.cell_value(rowNum, colNum) * 24
				hourInt = int(math.floor(hourFloat))

				self.hours.append({
					'h': hourInt,
					'm': int((hourFloat - hourInt) * 60)
				})

			colNum += 1

###########################################################
# iCalendar class
###########################################################
# TODO: add method
# TODO: initiate ics on init method
class iCalendar:
	def __init__(self, file):
		self.file = file
		self.cal = Calendar()

		self.cal.add('version', '2.0')
		self.cal.add('prodid', '-//Leeroy Brun - HEIG-VD//mxm.dk//')


	def add(self, summary, dateStart, dateEnd):
		print 'add date to ics'

	def save(self):
		print 'save ics'

# MAIN
timetable = Timetable(sys.argv[1])
timetable.parse()