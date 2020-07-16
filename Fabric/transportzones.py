import requests
import urllib3
import sys
import xlwt
import os

from xlwt import Workbook

from vmware.vapi.lib import connect
from vmware.vapi.stdlib.client.factories import StubConfigurationFactory
from com.vmware.nsx_client import TransportZones
from com.vmware.nsx.model_client import TransportZone
from vmware.vapi.security.user_password import \
        create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<OUPUT DIRECTORY>")

# Setup excel workbook and worksheets 
ls_wkbk = Workbook()  
sheet1 = ls_wkbk.add_sheet('Transport Zones')

#Set Excel Styling
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')
style_alignleft = xlwt.easyxf('font: colour black, bold False; align: horiz left')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 30
columnB = sheet1.col(1)
columnB.width = 256 * 40
columnC = sheet1.col(2)
columnC.width = 256 * 40
columnD = sheet1.col(3)
columnD.width = 256 * 20
columnE = sheet1.col(4)
columnE.width = 256 * 40
columnF = sheet1.col(5)
columnF.width = 256 * 22
columnG = sheet1.col(6)
columnG.width = 256 * 25
columnH = sheet1.col(7)
columnH.width = 256 * 25
columnI = sheet1.col(8)
columnI.width = 256 * 20
columnJ = sheet1.col(9)
columnJ.width = 256 * 25
columnJ = sheet1.col(10)
columnJ.width = 256 * 25

#Excel Column Headings
sheet1.write(0, 0, 'NAME', style_db_center)
sheet1.write(0, 1, 'DESCRIPTION', style_db_center)
sheet1.write(0, 2, 'ID', style_db_center)
sheet1.write(0, 3, 'RESOURCE TYPE', style_db_center)
sheet1.write(0, 4, 'HOST SWITCH ID', style_db_center)
sheet1.write(0, 5, 'HOST SWITCH MODE', style_db_center)
sheet1.write(0, 6, 'HOST SWITCH NAME', style_db_center)
sheet1.write(0, 7, 'HOST SWITCH IS DEFAULT', style_db_center)
sheet1.write(0, 8, 'IS NESTED NSX', style_db_center)
sheet1.write(0, 9, 'TRANSPORT TYPE', style_db_center)
sheet1.write(0, 10, 'UPLINK TEAMING POLICY NAME', style_db_center)

#Main Function
def main():
    session = requests.session()
    session.verify = False
    nsx_url = 'https://%s:%s' % ("<NSX MANAGER IP>", 443)
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_url)
    stub_config = StubConfigurationFactory.new_std_configuration(connector)
    security_context = create_user_password_security_context("<USERNAME>", "<PASSWORD>")
    connector.set_security_context(security_context)
    urllib3.disable_warnings()
    
    tz_list = []
    tz_svc = TransportZones(stub_config)
    tz_list = tz_svc.list()
    r = tz_list.results
    start_row = 1
    for i in r:
        tz = i.convert_to(TransportZone)
        sheet1.write(start_row, 0, tz.display_name)
        sheet1.write(start_row, 1, tz.description)
        sheet1.write(start_row, 2, tz.id)
        sheet1.write(start_row, 3, tz.resource_type)
        sheet1.write(start_row, 4, tz.host_switch_id)
        sheet1.write(start_row, 5, tz.host_switch_mode)
        sheet1.write(start_row, 6, tz.host_switch_name)
        sheet1.write(start_row, 7, tz.is_default)
        sheet1.write(start_row, 8, tz.nested_nsx)
        sheet1.write(start_row, 9, tz.transport_type)
        sheet1.write(start_row, 10,tz.uplink_teaming_policy_names)
        start_row += 1
    
    ls_wkbk.save('Transport Zones.xls')
    
if __name__ == "__main__":
    main()
