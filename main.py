#!/usr/bin/python

import sys
from Timetable import Timetable

if len(sys.argv) < 2:
	print "Usage : ", sys.argv[0], " sourceFile [filter]"
	sys.exit()
else:
	sourceFile = sys.argv[1]

if len(sys.argv) < 3:
	catFilter = ''
else:
	catFilter = sys.argv[2]

# MAIN
timetable = Timetable(sourceFile, catFilter)
timetable.parse()