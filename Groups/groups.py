import requests
import urllib3
import sys
import xlwt
import os

from xlwt import Workbook

from vmware.vapi.lib import connect
from vmware.vapi.stdlib.client.factories import StubConfigurationFactory
from com.vmware.nsx_policy.infra.domains_client import Groups
from com.vmware.nsx_policy.infra.domains.groups.members_client import IpAddresses  
from com.vmware.nsx_policy.infra.domains.groups.members_client import SegmentPorts 
from com.vmware.nsx_policy.infra.domains.groups.members_client import Segments  
from com.vmware.nsx_policy.infra.domains.groups.members_client import VirtualMachines  

from com.vmware.nsx_policy.model_client import PolicyGroupIPMembersListResult

from vmware.vapi.security.user_password import \
    create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
groups_wkbk = Workbook()  
sheet1 = groups_wkbk.add_sheet('Groups')

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
columnC.width = 256 * 30
columnD = sheet1.col(3)
columnD.width = 256 * 30
columnE = sheet1.col(4)
columnE.width = 256 * 30
columnF = sheet1.col(5)
columnF.width = 256 * 30
columnG = sheet1.col(6)
columnG.width = 256 * 50

#Excel Column Headings
sheet1.write(0, 0, 'GROUP NAME', style_db_center)
sheet1.write(0, 1, 'TAGS', style_db_center)
sheet1.write(0, 2, 'SCOPE', style_db_center)
sheet1.write(0, 3, 'IP ADDRESSES', style_db_center)
sheet1.write(0, 4, 'VIRTUAL MACHINES', style_db_center)
sheet1.write(0, 5, 'SEGMENTS', style_db_center)
sheet1.write(0, 6, 'SEGMENT PORTS', style_db_center)


def main():
    session = requests.session()
    session.verify = False
    nsx_url = 'https://%s:%s' % ("<NSX-T MANAGER FQDN>", 443)
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_url)
    stub_config = StubConfigurationFactory.new_std_configuration(connector)
    security_context = create_user_password_security_context("<USERNAME>", "<PASSWORD>")
    connector.set_security_context(security_context)
    urllib3.disable_warnings()
    domain_id = 'default'
    group_list = []
    group_svc = Groups(stub_config)
    group_list = group_svc.list(domain_id)
    # print(group_list)
    x = len(group_list.results)
    start_row = 1
    for i in range(0,x):
        # Extract Group ID for each group
        grp_id = group_list.results[i].id
        sheet1.write(start_row, 0, grp_id)

        # Extract Tags for each group if exist
        # Bypass system groups for LB
        if 'NLB.PoolLB' in grp_id or 'NLB.VIP' in grp_id: 
            pass
        elif group_list.results[i].tags:
            result = group_list.results[i].tags
            x = len(result)
            tag_list = []
            scope_list = []
            for i in range(0,x):
                tag_list.append(result[i].tag)
                scope_list.append(result[i].scope)
            sheet1.write(start_row, 1, ', '.join(tag_list), style_alignleft)    
            sheet1.write(start_row, 2, ', '.join(scope_list), style_alignleft) 
        
        # Bypass system groups for LB
        if 'NLB.PoolLB' in grp_id or 'NLB.VIP' in grp_id:  
            pass
        else:     
            # Create IP Address List for each group
            iplist = []
            ipsvc = IpAddresses(stub_config)
            iplist = ipsvc.list(domain_id, grp_id)
            iprc = len(iplist.results)
            iplist1 = []
            for i in range(0,iprc):
                iplist1.append(iplist.results[i])
            sheet1.write(start_row, 3, ', '.join(iplist1), style_alignleft)
            
            # Create Virtual Machine List for each group
            vmlist = []
            vmsvc = VirtualMachines(stub_config)
            vmlist = vmsvc.list(domain_id, grp_id)
            vmrc = vmlist.result_count
            vmlist1 = []
            for i in range(0,vmrc):
                vmlist1.append(vmlist.results[i].display_name)
            sheet1.write(start_row, 4, ', '.join(vmlist1), style_alignleft)

            # Create Segment List for each group
            sgmntlist = []
            sgmntsvc = Segments(stub_config)
            sgmntlist = sgmntsvc.list(domain_id, grp_id)
            sgmntrc = sgmntlist.result_count
            sgmntlist1 = []
            for i in range(0,sgmntrc):
                sgmntlist1.append(sgmntlist.results[i].display_name)
            sheet1.write(start_row, 5, ', '.join(sgmntlist1), style_alignleft)

            # Create Segment Port/vNIC List for each group
            sgmntprtlist = []
            sgmntprtsvc = SegmentPorts(stub_config)
            sgmntprtlist = sgmntprtsvc.list(domain_id, grp_id)
            sgmntprtrc = sgmntprtlist.result_count
            sgmntprtlist1 = []
            for i in range(0,sgmntprtrc):
                sgmntprtlist1.append(sgmntprtlist.results[i].display_name)
            sheet1.write(start_row, 6, ', '.join(sgmntprtlist1), style_alignleft)

        start_row +=1
    
    groups_wkbk.save('Groups.xls')

if __name__ == "__main__":
    main()
