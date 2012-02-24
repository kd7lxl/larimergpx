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
extralinesep = '' # change to '<br>' for Nuvi 350

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
    f.write("""<?xml version="1.0" encoding="UTF-8"?>\r\n<gpx xmlns="http://www.topografix.com/GPX/1/1" creator="larimergpx python" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd">\r\n""")

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
            # lat or lon is not a number, skip
            continue
        page = row[1].value or ''
        grid = row[2].value or ''
        sarutm = row[3].value
        topo = row[12].value
        county = row[13].value
        state = row[14].value
        
        try:
            if format is GPX:
                f.write("""  <wpt lat="%(latitude)s" lon="%(longitude)s">\r\n""" \
                        """    <ele>%(elevation)d</ele>\r\n""" \
                        """    <name>%(name)s</name>\r\n""" \
                       #"""    <cmt>%(name)s</cmt>\r\n""" \
                        """    <desc>""" \
                        """Gazetteer: %(page)s %(grid)s%(extralinesep)s\r\n""" \
                        """SAR UTM: %(sarutm)s%(extralinesep)s\r\n""" \
                        """Topo Map: %(topo)s%(extralinesep)s\r\n""" \
                        """%(county)s, %(state)s</desc>\r\n""" \
                        """  </wpt>\r\n""" % locals())
            elif format is CSV:
                f.write("""%(latitude)s,%(longitude)s,%(name)s,""" \
                        """"Gazetteer: %(page)s %(grid)s%(extralinesep)s\r\n""" \
                        """SAR UTM: %(sarutm)s%(extralinesep)s\r\n""" \
                        """Topo Map: %(topo)s%(extralinesep)s\r\n""" \
                        """%(county)s, %(state)s"\r\n""" % locals())
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
