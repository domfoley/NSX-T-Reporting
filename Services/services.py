import xlwt 
import os
import urllib.request
import urllib.parse
import urllib3
import ssl
import json
import requests
import csv 

from pprint import pprint
from xlwt import Workbook 

# Create auto-login for NSX (main script will require input)
ssl._create_default_https_context = ssl._create_unverified_context
s = requests.Session()
s.verify = False
s.auth = ('<username>', '<Password>')
nsx_mgr = '<NSX Manager fqdn>'
urllib3.disable_warnings()

# Auotmatic output directory for Excel Workbook
os.chdir("<Output directory for excel file>")

# Setup excel workbook and worksheets 
dfw_wkbk = Workbook()  
sheet1 = dfw_wkbk.add_sheet('NSX-T Services')
style_wrap = xlwt.easyxf('align: wrap True')
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 60
columnB = sheet1.col(1)
columnB.width = 256 * 60
columnC = sheet1.col(2)
columnC.width = 256 * 20
columnD = sheet1.col(3)
columnD.width = 256 * 60
columnE = sheet1.col(4)
columnE.width = 256 * 30
columnF = sheet1.col(5)
columnF.width = 256 * 20

style_wrap = xlwt.easyxf('alignment: wrap True')

services_upath = '/policy/api/v1/infra/services'
services_json = s.get(nsx_mgr + services_upath).json()

service_count = (services_json["result_count"])

sheet1.write(0, 0, 'Service Name', style_db_center)
sheet1.write(0, 1, 'Service Entries', style_db_center)
sheet1.write(0, 2, 'Service Type', style_db_center)
sheet1.write(0, 3, 'Port # / Additional Properties', style_db_center)
sheet1.write(0, 4, 'Tags', style_db_center)
sheet1.write(0, 5, 'Scope', style_db_center)

start_row = 1

for i in range(1,service_count):
    sheet1.write(start_row, 0, (services_json["results"][i]["display_name"]))
    svc_entries = (services_json["results"][i]["service_entries"])
    if "tags" in services_json["results"][i]:
            tag_length = (len(services_json["results"][i]["tags"]))
            tag_list = []
            scope_list = []
            for t in range(0,tag_length):
                tag_list.append(services_json["results"][i]["tags"][t]["tag"])
                scope_list.append(services_json["results"][i]["tags"][t]["scope"])
            sheet1.write(start_row, 4, ', '.join(tag_list))
            sheet1.write(start_row, 5, ', '.join(scope_list), style_wrap)
    for se in range(0,len(svc_entries)):
        sheet1.write(start_row, 1, (services_json["results"][i]["service_entries"][se]["id"]))
        if "l4_protocol" in services_json["results"][i]["service_entries"][se]:
            sheet1.write(start_row, 2, (services_json["results"][i]["service_entries"][se]["l4_protocol"]))
            d_ports = ",  "
            s = (services_json["results"][i]["service_entries"][se]["destination_ports"])
            sheet1.write(start_row, 3, (d_ports.join(s)))
        elif "protocol" in services_json["results"][i]["service_entries"][se]:
            prot = (services_json["results"][i]["service_entries"][se])
            sheet1.write(start_row, 2, (prot["protocol"]))
            if "icmp_type" in prot and "icmp_code" in prot:
                i_type = str(prot["icmp_type"])
                i_code = str(prot["icmp_code"])
                sheet1.write(start_row, 3, ('ICMP TYPE: '+i_type, '    ','ICMP CODE: '+i_code))
        elif "alg" in services_json["results"][i]["service_entries"][se]:
            sheet1.write(start_row, 2, (services_json["results"][i]["service_entries"][se]["alg"]))
            sheet1.write(start_row, 3, (services_json["results"][i]["service_entries"][se]["destination_ports"]))
        elif "protocol_number" in services_json["results"][i]["service_entries"][se]:
            pn = str(services_json["results"][i]["service_entries"][se]["protocol_number"])
            sheet1.write(start_row, 2, ('Protocol Number: ',(pn)))
        elif "ether_type" in services_json["results"][i]["service_entries"][se]:
            e_type = str(services_json["results"][i]["service_entries"][se]["ether_type"])
            sheet1.write(start_row, 2, ('Ether Type: ',(e_type)))
        else:
            sheet1.write(start_row, 2, ('IGMP'))
        start_row+=1

dfw_wkbk.save('services.xls') 

