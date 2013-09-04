#!/usr/bin/python
# coding=utf-8

import datetime
import math

import xlrd
from iCalendar import iCalendar

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

		self.ics = iCalendar('events.ics')

	#------------------------------------------------------
	# Main methods
	#------------------------------------------------------

	# Function used to parse a timetable xls file
	def parse(self):
		self.wb = xlrd.open_workbook(self.file, formatting_info=True)
		self.sh = self.wb.sheet_by_index(0)

		self.parseHours(HOURS_ROW)

		for rowNum in range(START_ROW, self.sh.nrows, WEEK_ROWS):
			self.parseWeek(rowNum)

		self.ics.save()

	# Parse rows of XLS file (which in fact is a week)
	def parseWeek(self, rowNum):
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
		for dayStartCol in range(START_COL, sh.row_len(rowNum), nbColsPerDay):
			if dayNum >= NB_DAYS:
				break;

			dayDate = weekStartDate + datetime.timedelta(days=dayNum)

			self.parseDay(rowNum, dayStartCol, dayDate)

			dayNum += 1

	def parseDay(self, rowNum, startColNum, dayDate):
		# Check if first day cell has color of "vacation" day
		if self.cellTypeByColor(rowNum, startColNum) == 'free':
			self.ics.add(dayDate, dayDate + datetime.timedelta(days=1), "HEIG-VD: Congé")
			return;

		# Ignore first col of day (contains DD.MM)
		startColNum += 1

		# Parse all courses of the day
		courseCol = startColNum
		while courseCol < startColNum+len(self.hours):
			# Empty cell ? Skip it.
			if(self.isCellEmpty(rowNum, courseCol) == True):
				courseCol += 1
			# Cell not empty, it's a course -> parse it and get the next col to check (last col of parsed course + 1)
			else:
				courseCol = self.parseCourse(rowNum, courseCol, dayDate, startColNum)
		

	def parseCourse(self, rowNum, colNum, dayDate, dayStartCol):
		maxDayCol = dayStartCol+len(self.hours)

		course = {}

		# Determine the course type
		cellTypeByColor = self.cellTypeByColor(rowNum, colNum)
		if cellTypeByColor == 'free':
			course['type'] = 'Congé'
		elif cellTypeByColor == 'test':
			course['type'] = 'Test'
		elif cellTypeByColor == 'exam':
			course['type'] = 'Examen'
		elif cellTypeByColor == 'test_exam':
			course['type'] = 'Test & Examen'
		else:
			course['type'] = 'Cours'

		# Find start datetime
		startTimeInDay = colNum - dayStartCol
		startTime = self.hours[startTimeInDay]
		course['startDate'] = dayDate.replace(hour=startTime['h'], minute=startTime['m'])

		# Get course data (name, teacher, location)
		course['name'] = self.sh.cell_value(rowNum, colNum)
		course['teacher'] = self.sh.cell_value(rowNum+1, colNum)
		course['location'] = self.sh.cell_value(rowNum+2, colNum)

		# Loop over next cells to find the end of the course
		colNum += 1
		endFound = False
		while endFound != True and colNum < maxDayCol:
			cellXf = self.cellXf(rowNum, colNum)

			if cellXf.border.right_line_style == 7:
				endFound = True
			else:
				colNum += 1

		# End found, use it
		if endFound == True:
			print rowNum, ":", colNum
			lastPeriodInDay = colNum - dayStartCol
		# End not found, the end is the last time of the day
		else:
			lastPeriodInDay = len(self.hours)-1

		# Process end date
		endTime = self.hours[lastPeriodInDay]
		course['endDate'] = dayDate.replace(hour=endTime['h'], minute=endTime['m'])
		# Need to add 45 min to the end date, because the time got before is the last period's start
		course['endDate'] = course['endDate'] + datetime.timedelta(minutes=45)

		tmp = {
			'summary': "HEIG-VD: "+ course['name'] +" ("+ course['type'] +")",
			'teacher': "Professeur : "+ course['teacher']
		} 

		self.ics.add(course['startDate'], course['endDate'], tmp['summary'], course['location'], tmp['teacher'])

		# Return the next cell to be processed by the parseDay loop
		return colNum+1

	# Get start hours of courses
	def parseHours(self, rowNum):
		sh = self.sh

		# 0 = not run, 1 = first empty cell met, 2 = second empty step met (STOP)
		emptyCells = 0
		colNum = START_COL

		self.hours = []

		while emptyCells != 2:
			if self.isCellEmpty(rowNum, colNum) == True:
				emptyCells += 1
			else:
				hourFloat = sh.cell_value(rowNum, colNum) * 24
				hourInt = int(math.floor(hourFloat))

				self.hours.append({
					'h': hourInt,
					'm': int((hourFloat - hourInt) * 60)
				})

			colNum += 1

	#------------------------------------------------------
	# Utils methods
	#------------------------------------------------------

	def cellXf(self, row, col):
		xfIndex = self.sh.cell_xf_index(row, col)
		return self.wb.xf_list[xfIndex]

	def isCellEmpty(self, row, col):
		cellType = self.sh.cell_type(row, col)
		if cellType == 0 or cellType == 6 or self.sh.cell_value(row, col) == ' ':
			return True
		else:
			return False

	def cellTypeByColor(self, row, col):
		cellColor = str(self.cellXf(row, col).background.pattern_colour_index)

		knownCellColors = {
			'47': 'test',
			'41': 'exam',
			'46': 'test_exam',
			'22': 'free'
		}

		if cellColor in knownCellColors:
			return knownCellColors[cellColor]
		else:
			return 'unknown'