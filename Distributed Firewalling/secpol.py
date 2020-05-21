import requests
import urllib3
import urllib.request
import urllib.parse
import sys
import ssl
import json
import xlwt
import os

from xlwt import Workbook

# Output directory for Excel Workbook
os.chdir("<OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
secpol_wkbk = Workbook()  
sheet1 = secpol_wkbk.add_sheet('Security Policies')

#Set Excel Styling
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')
style_alignleft = xlwt.easyxf('font: colour black, bold False; align: horiz left, wrap True')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 30
columnB = sheet1.col(1)
columnB.width = 256 * 30
columnC = sheet1.col(2)
columnC.width = 256 * 60
columnD = sheet1.col(3)
columnD.width = 256 * 20
columnE = sheet1.col(4)
columnE.width = 256 * 30
columnF = sheet1.col(5)
columnF.width = 256 * 20

#Excel Column Headings
sheet1.write(0, 0, 'SECURITY POLICY ID', style_db_center)
sheet1.write(0, 1, 'SECURITY POLICY NAME', style_db_center)
sheet1.write(0, 2, 'NSX POLICY PATH', style_db_center)
sheet1.write(0, 3, 'SEQUENCE NUMBER', style_db_center)
sheet1.write(0, 4, 'CATEGORY', style_db_center)
sheet1.write(0, 5, 'IS STATEFUL', style_db_center)

def main():
    ssl._create_default_https_context = ssl._create_unverified_context
    s = requests.Session()
    s.verify = False
    s.auth = ('<USERNAME>', '<PASSWORD>')
    nsx_mgr = 'https://<NSX-T MANAGER FQDN>'
    urllib3.disable_warnings()
    policies_upath = '/policy/api/v1/infra/domains/default/security-policies'
    policies_json = s.get(nsx_mgr + policies_upath).json()
    x = len(policies_json["results"])
    start_row = 1
    for i in range(0,x):
        sheet1.write(start_row, 0, policies_json["results"][i]["id"], style_alignleft)
        sheet1.write(start_row, 1, policies_json["results"][i]["display_name"], style_alignleft)
        sheet1.write(start_row, 2, policies_json["results"][i]["path"], style_alignleft)
        sheet1.write(start_row, 3, policies_json["results"][i]["sequence_number"], style_alignleft)      
        sheet1.write(start_row, 4, policies_json["results"][i]["category"], style_alignleft)
        sheet1.write(start_row, 5, policies_json["results"][i]["stateful"], style_alignleft)
        start_row +=1
    secpol_wkbk.save('Security Policies.xls')

if __name__ == "__main__":
    main()
