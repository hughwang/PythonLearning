from os import path
from pykml import parser

kml_file = path.join( 'sample_output.kml')

with open(kml_file) as f:
    doc = parser.parse(f).getroot()

    print doc.Document.Folder.Placemark.name
    print doc.Document.Folder.Placemark.Polygon.outerBoundaryIs.LinearRing.coordinates


