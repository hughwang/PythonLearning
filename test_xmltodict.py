#!/usr/bin/env python
# -*- coding: utf-8 -*-

# read the configuration file
import os
import re
import urllib2
import json
import gspread
from oauth2client.client import SignedJwtAssertionCredentials
from os import path

import xmltodict
outfile_name = 'DC_Posters.kml'
with open(outfile_name) as filehandle:
    doc = xmltodict.parse(filehandle.read())
placesmarks = doc['kml']['Document']['Placemark']
for item in placesmarks:
    name = item['name']
    coordinates = item['Polygon']['outerBoundaryIs']['LinearRing']['coordinates']
    print name, coordinates
    



i=1