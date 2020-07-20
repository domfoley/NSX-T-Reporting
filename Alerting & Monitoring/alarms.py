import requests
import urllib3
import sys
import xlwt
import os
import datetime

from xlwt import Workbook

from vmware.vapi.lib import connect
from vmware.vapi.bindings.stub import VapiInterface
from vmware.vapi.bindings.stub import StubConfiguration
from vmware.vapi.stdlib.client.factories import StubConfigurationFactory
from com.vmware.nsx_client import Alarms
from vmware.vapi.security.user_password import \
    create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
ls_wkbk = Workbook()  
sheet1 = ls_wkbk.add_sheet('NSX Alarms')

#Set Excel Styling
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')
style_alignleft = xlwt.easyxf('font: colour black, bold False; align: horiz left, wrap True')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 20
columnB = sheet1.col(1)
columnB.width = 256 * 40
columnC = sheet1.col(2)
columnC.width = 256 * 30
columnD = sheet1.col(3)
columnD.width = 256 * 20
columnE = sheet1.col(4)
columnE.width = 256 * 30
columnF = sheet1.col(5)
columnF.width = 256 * 15
columnG = sheet1.col(6)
columnG.width = 256 * 20
columnH = sheet1.col(7)
columnH.width = 256 * 20
columnI = sheet1.col(8)
columnI.width = 256 * 30
columnJ = sheet1.col(9)
columnJ.width = 256 * 60

#Excel Column Headings
sheet1.write(0, 0, 'Feature', style_db_center)
sheet1.write(0, 1, 'Event Type', style_db_center)
sheet1.write(0, 2, 'Reporting Node', style_db_center)
sheet1.write(0, 3, 'Node Resource Type', style_db_center)
sheet1.write(0, 4, 'Entity Name', style_db_center)
sheet1.write(0, 5, 'Severity', style_db_center)
sheet1.write(0, 6, 'Last Reported Time', style_db_center)
sheet1.write(0, 7, 'Status', style_db_center)
sheet1.write(0, 8, 'Description', style_db_center)
sheet1.write(0, 9, 'Recommended Action', style_db_center)

def main():
    session = requests.session()
    session.verify = False
    nsx_mgr = 'https://%s' % ("1<NSXT MGR IP / FQDN>")
    nsx_auth = ('<USERNAME>', '<PASSWORD>')
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_mgr)
    stub_config = StubConfigurationFactory.new_std_configuration(connector)
    security_context = create_user_password_security_context(nsx_auth[0], nsx_auth[1])
    connector.set_security_context(security_context)
    urllib3.disable_warnings()

    tn_url = '/api/v1/transport-nodes'
    tn_json = session.get(nsx_mgr + str(tn_url), auth=nsx_auth, verify=session.verify).json()
    t_nodes = len(tn_json["results"])

    mgr_url = '/api/v1/cluster/nodes'
    mgr_json = session.get(nsx_mgr + str(mgr_url), auth=nsx_auth, verify=session.verify).json()
    mgr_nodes = len(mgr_json["results"])

    node_dict = {}
    
    for res in range(0,t_nodes):
        node = tn_json["results"][res]
        node_dict.update({node["node_id"]:node["display_name"]})
    
    for res in range(0,mgr_nodes):
        node = mgr_json["results"][res]
        node_dict.update({node["id"]:node["display_name"]})
    
    alarms_list = []
    alarms_svc = Alarms(stub_config)
    alarms_list = alarms_svc.list()

    x = (len(alarms_list.results))
    y = alarms_list.results
    start_row = 1
    
    for i in range(x):
        sheet1.write(start_row, 0, y[i].feature_name, style_alignleft)
        sheet1.write(start_row, 1, y[i].event_type, style_alignleft)

        for key, value in node_dict.items():
            if key == y[i].node_id:
                sheet1.write(start_row, 2, value, style_alignleft)
            # else:
            #     print(y[i].node_id)

        sheet1.write(start_row, 3, y[i].node_resource_type, style_alignleft)
        sheet1.write(start_row, 4, y[i].entity_id, style_alignleft)
        sheet1.write(start_row, 5, y[i].severity, style_alignleft)

        lrt = y[i].last_reported_time
        dtt = datetime.datetime.fromtimestamp(float(lrt/1000)).strftime('%Y-%m-%d %H:%M:%S')
        
        sheet1.write(start_row, 6, dtt, style_alignleft)
        sheet1.write(start_row, 7, y[i].status, style_alignleft)
        sheet1.write(start_row, 8, y[i].description, style_alignleft)
        sheet1.write(start_row, 9, y[i].recommended_action, style_alignleft)

        start_row +=1
    
    ls_wkbk.save('Alarms.xls')

if __name__ == "__main__":
    main()
