#!/usr/bin/python

import sys
import datetime
import uuid

from icalendar import Calendar, Event, LocalTimezone

# LINKS :
#  - Save to .ics : http://icalendar.readthedocs.org/en/latest/
#       http://icalendar.readthedocs.org/en/latest/examples.html
#       http://icalendar.readthedocs.org/en/latest/examples.html#more-documentation
#       

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
		self.cal.add('prodid', '-//Leeroy Brun - HEIG-VD//Timetable Parser//')

	def add(self, dateStart, dateEnd, summary, location='', description='', categories=''):
		event = Event()

		print categories

		event.add('summary', summary)
		event.add('location', location)
		event.add('categories', categories, encode=0)
		event.add('description', description)
		event.add('dtstart', dateStart)
		event.add('dtend', dateEnd)
		event.add('dtstamp', datetime.datetime.now())
		event['uid'] = uuid.uuid4()
		event.add('priority', 5)

		self.cal.add_component(event)

	def save(self):
		f = open(self.file, 'wb')
		f.write(self.cal.to_ical())
		f.close()
