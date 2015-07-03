import json
import gspread
from oauth2client.client import SignedJwtAssertionCredentials

json_key = json.load(open('my_credentials_get_From_google_site_need_replace.json'))
scope = ['https://spreadsheets.google.com/feeds']

credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'], scope)

gc = gspread.authorize(credentials)

excel_doc = gc.open("DC Offices Assignment")
#worksheet = excel_doc.worksheet("OfficesMaps")
worksheet = excel_doc.worksheet("Test")

values_list = worksheet.row_values(1)
print values_list

worksheet.update_cell(5, 2, 'Hello')




"""
please refer to following article "Access google sheets in python using gspread"
http://www.indjango.com/access-google-sheets-in-python-using-gspread/

http://gspread.readthedocs.org/en/latest/oauth2.html
https://github.com/burnash/gspread
"""
