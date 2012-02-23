#!/usr/bin/env python
"""
Converts xlsx files based on LarimerCountyGNISExtractGazetteer2008-12-02_rv4.xlsx
to a gpx waypoint file.


Copyright (C) 2012  Tom Hayward

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

import sys

from openpyxl.reader.excel import load_workbook

FEET_TO_METERS = 0.3048

# input xlsx
try:
    wb = load_workbook(filename=sys.argv[1])
except IndexError:
    print >> sys.stderr, "Usage: %s XLSX_FILE [GPX_FILE]" % sys.argv[0]
    sys.exit(1)
sheet = wb.get_sheet_by_name(name='GNIS data extract on 2 Dec 2008')

# output gpx
try:
    f = open(sys.argv[2], 'w')
except IndexError:
    f = sys.stdout

f.write("""<?xml version="1.0" encoding="UTF-8"?>
<gpx version="1.0">
""")

try:
    for row in sheet.rows[4:]:
        name = row[0].value
        try:
            elevation = row[7].value * FEET_TO_METERS
        except TypeError:
            # elevation is not a number, skip
            continue
        latitude = row[8].value
        longitude = row[9].value
        if type(longitude) is not float or type(latitude) is not float:
            # lat or long is not a number, skip
            continue
        
        try:
            f.write("""	<wpt lat="%s" lon="%s">
		<ele>%d</ele>
		<name>%s</name>
	</wpt>
""" % (latitude, longitude, elevation, name))
        except UnicodeEncodeError:
            pass
except KeyboardInterrupt:
    print >> sys.stderr, "Aborting."
finally:
    f.write("</gpx>\n")