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

from vmware.vapi.lib import connect
from vmware.vapi.stdlib.client.factories import StubConfigurationFactory
from vmware.vapi.bindings.stub import StubConfiguration
from com.vmware.nsx.fabric_client import DiscoveredNodes
from vmware.vapi.security.user_password import \
        create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
d_node_wkbk = Workbook()  
sheet1 = d_node_wkbk.add_sheet('Discovered Nodes')

#Set Excel Styling
style_db_center = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz center')
style_alignleft = xlwt.easyxf('font: colour black, bold True; align: horiz left, wrap True')

#Setup Column widths
columnA = sheet1.col(0)
columnA.width = 256 * 30
columnB = sheet1.col(1)
columnB.width = 256 * 20
columnC = sheet1.col(2)
columnC.width = 256 * 20
columnD = sheet1.col(3)
columnD.width = 256 * 20
columnE = sheet1.col(4)
columnE.width = 256 * 20
columnF = sheet1.col(5)
columnF.width = 256 * 35
columnG = sheet1.col(6)
columnG.width = 256 * 15
columnH = sheet1.col(7)
columnH.width = 256 * 20
columnI = sheet1.col(8)
columnI.width = 256 * 15
columnJ = sheet1.col(9)
columnJ.width = 256 * 40
columnK = sheet1.col(10)
columnK.width = 256 * 15
columnL = sheet1.col(11)
columnL.width = 256 * 20
columnM = sheet1.col(12)
columnM.width = 256 * 15
columnN = sheet1.col(13)
columnN.width = 256 * 20
columnO = sheet1.col(14)
columnO.width = 256 * 20
columnP = sheet1.col(15)
columnP.width = 256 * 50
columnQ = sheet1.col(16)
columnQ.width = 256 * 20
columnR = sheet1.col(17)
columnR.width = 256 * 25
columnS = sheet1.col(18)
columnS.width = 256 * 25
columnT = sheet1.col(19)
columnT.width = 256 * 20
columnU = sheet1.col(20)
columnU.width = 256 * 20
columnV = sheet1.col(21)
columnV.width = 256 * 20

#Excel Column Headings
sheet1.write(0, 0, 'Number of discovered nodes: ', style_alignleft)
sheet1.write(2, 0, 'Display name', style_db_center)
sheet1.write(2, 1, 'OS Type', style_db_center)
sheet1.write(2, 2, 'OS Version', style_db_center)
sheet1.write(2, 3, 'Node Type', style_db_center)
sheet1.write(2, 4, 'Hostname', style_db_center)
sheet1.write(2, 5, 'Full Name', style_db_center)
sheet1.write(2, 6, 'Management IP', style_db_center)
sheet1.write(2, 7, 'Domain name', style_db_center)
sheet1.write(2, 8, 'DNS', style_db_center)
sheet1.write(2, 9, 'UUID', style_db_center)
sheet1.write(2, 10, 'Powerstate', style_db_center)
sheet1.write(2, 11, 'In Maintenance Mode', style_db_center)
sheet1.write(2, 12, 'Build', style_db_center)
sheet1.write(2, 13, 'Vendor', style_db_center)
sheet1.write(2, 14, 'Model', style_db_center)
sheet1.write(2, 15, 'Serial Number', style_db_center)
sheet1.write(2, 16, 'Connection State', style_db_center)
sheet1.write(2, 17, 'Licensed Product Name', style_db_center)
sheet1.write(2, 18, 'Licensed Product Version', style_db_center)
sheet1.write(2, 19, 'Mgmt Server IP', style_db_center)
sheet1.write(2, 20, 'Lockdown Mode', style_db_center)
sheet1.write(2, 21, 'DAS Host State', style_db_center)


def main():
    session = requests.session()
    session.verify = False
    nsx_url = 'https://%s:%s' % ("10.29.13.111", 443)
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_url)
    stub_config = StubConfigurationFactory.new_std_configuration(connector)
    security_context = create_user_password_security_context("admin", "VMware1!vmware")
    connector.set_security_context(security_context)
    urllib3.disable_warnings()
    disc_node_list = []
    disc_node_svc = DiscoveredNodes(stub_config)
    disc_node_list = disc_node_svc.list()
    discovered_nodes = disc_node_list.result_count
    sheet1.write(0, 1, discovered_nodes, style_alignleft)
    
    start_row = 3
    for i in range(discovered_nodes):
        sheet1.write(start_row, 0, disc_node_list.results[i].display_name)
        sheet1.write(start_row, 1, disc_node_list.results[i].os_type)
        sheet1.write(start_row, 2, disc_node_list.results[i].os_version)
        sheet1.write(start_row, 3, disc_node_list.results[i].node_type)
        
        origin_list = disc_node_list.results[i].origin_properties
        origin_dict = dict.fromkeys(origin_list)

        for key in origin_dict.keys():
            if key.key == 'hostName':
                sheet1.write(start_row, 4, key.value)
            if key.key == 'fullName':
                sheet1.write(start_row, 5, key.value)
            if key.key == 'managementIp':
                sheet1.write(start_row, 6, key.value)
            if key.key == 'domainName':
                sheet1.write(start_row, 7, key.value)
            if key.key == 'dnsConfigAddress':
                sheet1.write(start_row, 8, key.value)
            if key.key == 'uuid':
                sheet1.write(start_row, 9, key.value)
            if key.key == 'powerState':
                sheet1.write(start_row, 10, key.value)
            if key.key == 'inMaintenanceMode':
                sheet1.write(start_row, 11, key.value)
            if key.key == 'build':
                sheet1.write(start_row, 12, key.value)
            if key.key == 'vendor':
                sheet1.write(start_row, 13, key.value)
            if key.key == 'model':
                sheet1.write(start_row, 14, key.value)
            if key.key == 'serialNumber':
                sheet1.write(start_row, 15, key.value)
            if key.key == 'connectionState':
                sheet1.write(start_row, 16, key.value)
            if key.key == 'licenseProductName':
                sheet1.write(start_row, 17, key.value)
            if key.key == 'licenseProductVersion':
                sheet1.write(start_row, 18, key.value)
            if key.key == 'managementServerIp':
                sheet1.write(start_row, 19, key.value)
            if key.key == 'lockdownMode':
                sheet1.write(start_row, 20, key.value)
            if key.key == 'dasHostState':
                sheet1.write(start_row, 21, key.value)
            
        start_row +=1
    
    d_node_wkbk.save('Discovered Nodes.xls')

if __name__ == "__main__":
    main()
