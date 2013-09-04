#!/usr/bin/python

import sys
from Timetable import Timetable

# MAIN
timetable = Timetable(sys.argv[1])
timetable.parse()

print "col 3"
print timetable.sh.cell(9, 3).dump()

print "col 4"
print timetable.sh.cell(9, 4).dump()

print "col 5"
print timetable.sh.cell(9, 5).dump()

print "col 6"
print timetable.sh.cell(9, 6).dump()

print "col 7"
print timetable.sh.cell(9, 7).dump()

print "col 8"
print timetable.sh.cell(9, 8).dump()