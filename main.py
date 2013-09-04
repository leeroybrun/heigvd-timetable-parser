#!/usr/bin/python

import sys
from Timetable import Timetable

# MAIN
timetable = Timetable(sys.argv[1])
timetable.parse()