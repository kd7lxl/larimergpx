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

import codecs
import sys

from openpyxl.reader.excel import load_workbook

FEET_TO_METERS = 0.3048

success = 0
errors = []

class CSV: pass
class GPX: pass

# input xlsx
try:
    wb = load_workbook(filename=sys.argv[1])
except IndexError:
    print >> sys.stderr, "Usage: %s XLSX_FILE [GPX_FILE]" % sys.argv[0]
    sys.exit(1)
sheet = wb.get_sheet_by_name(name='GNIS data extract on 2 Dec 2008')

# output gpx
format = GPX
try:
    f = codecs.open(sys.argv[2], encoding='utf-8', mode='w')
    if sys.argv[2].endswith('csv'):
        format = CSV
except IndexError:
    f = sys.stdout

if format is GPX:
    f.write("""<?xml version="1.0" encoding="UTF-8"?>\r\n<gpx version="1.1" creator="larimergpx in python">
""")

try:
    for row in sheet.rows:
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
            if format is GPX:
                f.write("""  <wpt lat="%s" lon="%s">\r\n    <ele>%d</ele>\r\n    <name>%s</name>\r\n    <cmt>%s</cmt>\r\n    <desc>%s</desc>\r\n  </wpt>\r\n""" % (latitude, longitude, elevation, name, name, name))
            elif format is CSV:
                f.write("""%s,%s,%s\r\n""" % (latitude, longitude, name))
            success += 1
        except UnicodeEncodeError:
            errors.append("Can't print unicode character in row %s, skipping." % row[0].row)
            print >> sys.stderr, "Error:", errors[-1]
            continue
        if f == sys.stdout:
            f.flush()
except IOError:
    pass
except KeyboardInterrupt:
    print >> sys.stderr, "Aborting."
finally:
    try:
        if format is GPX:
            f.write("</gpx>")
        f.close()
    except IOError:
        pass

print >> sys.stderr, "\n%d waypoints exported.\n" % success
if errors:
    print >> sys.stderr, "Error summary:"
    for error in errors:
        print >> sys.stderr, '-', error
    print >> sys.stderr
    sys.exit(1)
sys.exit(0)
