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
from com.vmware.nsx_client import LogicalSwitches
from com.vmware.nsx.model_client import LogicalSwitch
from vmware.vapi.security.user_password import \
        create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<PATH OF OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
ls_wkbk = Workbook()  
sheet1 = ls_wkbk.add_sheet('Logical Switching')

#Set Excel Styling
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')
style_alignleft = xlwt.easyxf('font: colour black, bold False; align: horiz left')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 40
columnB = sheet1.col(1)
columnB.width = 256 * 15
columnC = sheet1.col(2)
columnC.width = 256 * 15
columnD = sheet1.col(3)
columnD.width = 256 * 40
columnE = sheet1.col(4)
columnE.width = 256 * 40
columnF = sheet1.col(5)
columnF.width = 256 * 15
columnG = sheet1.col(6)
columnG.width = 256 * 20
columnH = sheet1.col(7)
columnH.width = 256 * 20
columnI = sheet1.col(8)
columnI.width = 256 * 40
columnJ = sheet1.col(9)
columnJ.width = 256 * 30

#Excel Column Headings
sheet1.write(0, 0, 'LOGICAL SWITCH', style_db_center)
sheet1.write(0, 1, 'VNI', style_db_center)
sheet1.write(0, 2, 'VLAN', style_db_center)
sheet1.write(0, 3, 'TRANSPORT ZONE NAME', style_db_center)
sheet1.write(0, 4, 'TRANSPORT ZONE ID', style_db_center)
sheet1.write(0, 5, 'TZ TYPE', style_db_center)
sheet1.write(0, 6, 'REPLICATION MODE', style_db_center)
sheet1.write(0, 7, 'ADMIN STATE', style_db_center)
sheet1.write(0, 8, 'POLICY PATH', style_db_center)
sheet1.write(0, 9, 'SUBNET', style_db_center)

#Main Function
def main():
    session = requests.session()
    session.verify = False
    nsx_url = 'https://%s:%s' % ("<NSX-T MANAGER FQDN>", 443)
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_url)
    stub_config = StubConfigurationFactory.new_std_configuration(connector)
    security_context = create_user_password_security_context("<USERNAME>", "<PASSWORD>")
    connector.set_security_context(security_context)
    urllib3.disable_warnings()
    ls_list = []
    ls_svc = LogicalSwitches(stub_config)
    ls_list = ls_svc.list()
    # print(ls_list)
    tz_list = []
    tz_svc = TransportZones(stub_config)
    tz_list = tz_svc.list()
    start_row = 1
    for vs in ls_list.results:
        ls = vs.convert_to(LogicalSwitch)
        sheet1.write(start_row, 0, ls.display_name)
        sheet1.write(start_row, 1, ls.vni, style_alignleft)
        sheet1.write(start_row, 2, ls.vlan, style_alignleft)
        sheet1.write(start_row, 4, ls.transport_zone_id)
        sheet1.write(start_row, 6, ls.replication_mode)
        sheet1.write(start_row, 7, ls.admin_state)
        # print(ls.resource_type)
        newlist = []
        for i in range(len(ls.tags)):
            newlist.append(ls.tags[i].tag)
        sheet1.write(start_row, 8,(str(newlist[0])))
        if len(newlist) > 1:
            sheet1.write(start_row, 9,(str(newlist.pop())))
        x = len(tz_list.results)
        for i in range(0,x):
            if ls.transport_zone_id == tz_list.results[i].id:
                sheet1.write(start_row, 3, tz_list.results[i].display_name)
                sheet1.write(start_row, 5, tz_list.results[i].transport_type)
        start_row +=1
    ls_wkbk.save('Logical Switches.xls')
    
if __name__ == "__main__":
    main()
