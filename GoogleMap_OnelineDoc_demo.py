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
from shapely.geometry import Polygon
from shapely.geometry import Point
import json
from pprint import pprint
import pyodbc
import os
from django.template.loader import get_template
from django import template
from django.template import Context


# use MS Access DB to store Offices information
DBfile = 'K:\\Baltimore\\Company.accdb'

conn = pyodbc.connect('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+DBfile)
cursor = conn.cursor()


#global variables
streetsAssign={}
AllStreets={}
AllStreets['All'] = []
AllStreets['NotVisited'] = []
ProcessedLocations={}

list_length=0
id = 0

def main():
        global id
        global list_length
        # read the config file
        with open('config.json', "r") as data_file:
                config_data = json.loads(data_file.read())
        #pprint(config_data)

        worksheet =0

        # read the online excel doc
	# for your reference:  http://www.indjango.com/access-google-sheets-in-python-using-gspread/
        json_key = json.load(open('pmt002-9d0ee0287ed5.json'))
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'], scope)
        gc = gspread.authorize(credentials)
        google_doc_title = config_data['GOOGLE_DOC_SPREADSHEET_NAME']
        excel_doc = gc.open(google_doc_title)
        worksheet_name = config_data['GOOGLE_DOC_SPREADSHEET_TAB_NAME']
        worksheet = excel_doc.worksheet(worksheet_name)

        list_of_lists = worksheet.get_all_values()
        row_1 = list_of_lists[0]
        collum_name_dic = {}
        collum_count = 0
        for field in row_1:
            collum_name_dic[field] = collum_count
            collum_count = collum_count+1

        # go through other rows
        list_length = len(list_of_lists)
        for row in list_of_lists[1:]:
            this_id= GetRowCellByField(row,collum_name_dic,'id')
            if re.match('\d+$', this_id) and id < int(this_id) :
                id = int(this_id)
            Address= GetRowCellByField(row,collum_name_dic,'Address')
            StreetName = GetRowCellByField(row,collum_name_dic,'Location')
            AssignedTo = GetRowCellByField(row,collum_name_dic,'AssignedTo')
            MapURL  = GetRowCellByField(row,collum_name_dic,'MapURL')
            streetsAssign[StreetName] = AssignedTo

        for stateInfo in config_data['state_maps']:
                State=    stateInfo['State']
                Google_Map_KML = stateInfo['GOOGLE_MAP_KML_URL']

                print "Process State %s now:" % (State)
                ProcessForState(config_data,worksheet,State,Google_Map_KML,collum_name_dic)

def    ProcessForState(config_data,worksheet,State,Google_Map_KML,collum_name_dic):
        global streetsAssign
        global AllStreets
        global id
        global cursor
        global list_length,ProcessedLocations
        
        # query db, get all Office in this State into rows
        SQL = "SELECT ID, Address,City,State,Zip,Latitude,Longitude,Company,Source from DCOffices where Distance < 40 and  State in ('%s') order by State,Zip,ID " % (State)
        cursor.execute(SQL)
        rows = cursor.fetchall()

        # make request to download the KML file from google map
        url = Google_Map_KML
        outfile_name = '%s_Offices.kml'%(State);
        outfile = open(outfile_name,'w')
        urlfile = urllib2.urlopen(url)
        if urlfile:
            outfile.write(urlfile.read())
            outfile.close()
        else:
            print "download %s failed" % (url)


        #################################################
        # rewrite the All.txt kml document ,
        # color the arleady assiged Street with special color
        ##################################################

        average_all_lat = 0;
        average_all_long = 0;
        count_all=0;
        styles=[];
        streets={};


        with open(outfile_name) as filehandle:
                doc = xmltodict.parse(filehandle.read())
        placesmarks = doc['kml']['Document']['Placemark']
        Styles = doc['kml']['Document']['Style']
        for item in Styles:
                Style_Name = item['@id']
                if 'LineStyle' in item:
                        LineColor = item['LineStyle']['color']
                elif 'IconStyle' in item:
                        LineColor = item['IconStyle']['color']
                elif 'PolyStyle' in item:
                        LineColor = item['PolyStyle']['color']
        styles.append([Style_Name,LineColor])


        for Street in placesmarks:
                print Street['name']
                StreetName = Street['name']
                StreetName = StreetName.strip()
                StreetName = RemoveStrangeCharacter(StreetName)

                mapType=''
                coordinates=''
                if 'LineString' in  Street:
                        coordinates = Street['LineString']['coordinates']
                        mapType='line';
                elif ( 'Polygon' in Street):
                        coordinates = Street['Polygon']['outerBoundaryIs']['LinearRing']['coordinates']
                        mapType = 'shape'
                else:
                        coordinates = Street['Point']['coordinates']
                        mapType='point'
                shortDescription=''

                StreetStyle= Street['styleUrl'];
                poly_points=[]
                coordinates_list = coordinates.split(' ')
                offices_list=[]
                for coordinate_string in coordinates_list:
                        longitude,latitude,z=coordinate_string.split(',')
                        poly_points.append((float(latitude), float(longitude)))

                poly = Polygon(poly_points,)
                # find out all locations within this polygon

                count_number = 0
                this_city=''
                this_Address=''
                for row in rows: # cursors are iterable
                        location_string ="%s_%s"%(row.Latitude,row.Longitude)
                        if location_string in ProcessedLocations:
                                continue
                        point = Point(row.Latitude,row.Longitude)
                        if point.within(poly):
                                #print "Point %s,%s in the Polygon"%(row.Latitude,row.Longitude)
                                Office={}
                                Office['ID']=row.ID
                                Office['Address']=row.Address
                                Office['State']=row.State
                                Office['City']=row.City
                                Office['Zip']=row.Zip
                                Office['Latitude']=row.Latitude
                                Office['Longitude']=row.Longitude
                                Office['Company']=row.Company
                                Office['Source']=row.Source                    
                                offices_list.append(Office)
                                this_city = row.City
                                this_Address = row.Address
                                ProcessedLocations[location_string] = 1


                if not StreetName in streetsAssign:
                        id = id + 1
                        list_length = list_length + 1
                        AddStreetToSpread(config_data,worksheet,collum_name_dic,id, StreetName,list_length,this_city,State,this_Address)
                        streetsAssign[StreetName] = ""

                streetValue = streetsAssign[StreetName]
                if streetValue:
                        StreetStyle = "assigned_street"
                else:
                        StreetStyle = "un_assigned_street"

                mapurl =  config_data['WEB_URL_FOR_STORED_MAP']+StreetName + '.htm'
                AllStreets['All'].append({'StreetName':StreetName,
                                          'StreetStyle':StreetStyle,
                                          'MapType':mapType,
                                          'Polygon_Points':poly_points,
                                          'mapurl':mapurl,
                                          'Office_List':[]})
                AllStreets['NotVisited'].append({'StreetName':StreetName,
                                          'StreetStyle':StreetStyle,
                                          'MapType':mapType,
                                          'Polygon_Points':poly_points,
                                          'mapurl':mapurl,
                                          'Office_List':[]})

                if StreetName not in AllStreets:
                        AllStreets[StreetName] = []
                AllStreets[StreetName].append({'StreetName':StreetName,
                                          'StreetStyle':StreetStyle,
                                          'MapType':mapType,
                                          'Polygon_Points':poly_points,
                                          'mapurl':mapurl,
                                          'Office_List':offices_list})                        


        for StreetGroup in AllStreets:
                if ( StreetGroup == "NotVisited" ):
                        continue
                ProcessStreetGroup(config_data,StreetGroup)

def ProcessStreetGroup (config_data,StreetGroup):     
        output_folder = config_data['OUTPUT_FOLDER']
        output= output_folder +"/%s.htm" % (StreetGroup)

        if not os.path.exists(output_folder):
                os.makedirs(output_folder)
        t = get_template('polygon_template.htm')
        point_list = AllStreets[StreetGroup][0]['Polygon_Points']
        count=0
        center_latitude= 0
        center_longitude = 0
        if(StreetGroup == "All"):
                center_latitude =38.896172
                center_longitude =-77.055630
                roomsize = 12
        else:   
                for point in point_list:
                      center_latitude = center_latitude + point[0]
                      center_longitude = center_longitude + point[1]
                      count = count + 1
                center_latitude = center_latitude / count
                center_longitude = center_longitude / count
                roomsize = 16
        
        c = template.Context({'StreetList': AllStreets[StreetGroup],
                              'center_latitude':center_latitude,
                              'center_longitude':center_longitude,
                              'roomsize':roomsize,
                              'StreetGroup':StreetGroup,
                              
                              })
        outfile = open(output,"w")
        outfile.write(t.render(c))
        outfile.close


def GetRowCellByField(row,collum_name_dic,field_name):
        return row[collum_name_dic[field_name]]

def RemoveStrangeCharacter(s):
        s = re.sub(r"/", '_', s)
        s = re.sub(r"\s+", '_', s)
        s = re.sub(r"#", '_', s)
        s = re.sub(r"\*", '_', s)
        s = re.sub(r"\(", '_', s)
        s = re.sub(r"\)", '_', s)
        s = re.sub(r"'", '_', s)
        s = re.sub(r"&", '_and_', s)
        s = re.sub(r"\\", '_', s)
        return s

def AddStreetToSpread(config_data,worksheet,column_name_dict,id, StreetName,row_count,City,State,Address):
        new_row={}
        new_row['id']=id
        new_row['AssignedTo']=''
        new_row['Location']=StreetName
        new_row['City']=City.title()
        new_row['Address']=Address
        new_row['State']=State.upper()
        new_row['MapURL']="%s/%s.htm"%(config_data['WEB_URL_FOR_STORED_MAP'], StreetName)

        for key in new_row:
                col = column_name_dict[key]
                worksheet.update_cell(row_count, col+1, new_row[key])

if __name__ == "__main__":
        os.environ.setdefault("DJANGO_SETTINGS_MODULE", "settings")
        main()
