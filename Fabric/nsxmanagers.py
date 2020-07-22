import requests
import urllib3
import sys
import xlwt
import os
import datetime

from xlwt import Workbook

from vmware.vapi.lib import connect
from vmware.vapi.security.user_password import \
        create_user_password_security_context

# Output directory for Excel Workbook
os.chdir("<OUTPUT DIRECTORY>")

# Setup excel workbook and worksheets 
ls_wkbk = Workbook()  
summary = ls_wkbk.add_sheet('NSX-T Summary')

#Set Excel Styling
style_db_left = xlwt.easyxf('pattern: pattern solid, fore_colour blue_grey;'
                                'font: colour white, bold True; align: horiz left, vert centre')
style_db_left1 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                'font: colour white, bold True; align: horiz left, vert centre')
style_alignleft = xlwt.easyxf('font: colour black, bold False; align: horiz left, wrap True')
style_green = xlwt.easyxf('font: colour green, bold True; align: horiz left, wrap True')
style_red = xlwt.easyxf('font: colour red, bold True; align: horiz left, wrap True')

def main():
    session = requests.session()
    session.verify = False
    nsx_mgr = 'https://%s' % ("<NSX MGR IP / FQDN>")
    nsx_auth = ('<USERNAME>', '<PASSWORD>')
    connector = connect.get_requests_connector(session=session, msg_protocol='rest', url=nsx_mgr)
    security_context = create_user_password_security_context(nsx_auth[0], nsx_auth[1])
    connector.set_security_context(security_context)
    urllib3.disable_warnings()

    ########### SECTION FOR REPORTING ON NSX-T MANAGER CLUSTER ###########

    nsxclstr_url = '/api/v1/cluster/status'
    nsxclstr_json = session.get(nsx_mgr + str(nsxclstr_url), auth=nsx_auth, verify=session.verify).json()
    
    columnA = summary.col(0)
    columnA.width = 256 * 35
    columnB = summary.col(1)
    columnB.width = 256 * 65
    
    summary.write(0, 0, 'NSX-T Cluster ID', style_db_left)
    summary.write(1, 0, 'NSX-T Cluster Status', style_db_left)
    summary.write(2, 0, 'NSX-T Control Cluster Status', style_db_left)
    summary.write(3, 0, 'Overall NSX-T Cluster Status', style_db_left)
    summary.write(0, 1, nsxclstr_json['cluster_id'], style_alignleft)
    if nsxclstr_json['mgmt_cluster_status']['status'] == 'STABLE':
        summary.write(1, 1, nsxclstr_json['mgmt_cluster_status']['status'], style_green)
    else:
        summary.write(1, 1, nsxclstr_json['mgmt_cluster_status']['status'], style_red)
    if nsxclstr_json['control_cluster_status']['status'] == 'STABLE':
        summary.write(2, 1, nsxclstr_json['control_cluster_status']['status'], style_green)
    else:
        summary.write(2, 1, nsxclstr_json['control_cluster_status']['status'], style_red)
    if nsxclstr_json['detailed_cluster_status']['overall_status'] == 'STABLE':
        summary.write(3, 1, nsxclstr_json['detailed_cluster_status']['overall_status'], style_green)
    else:
        summary.write(3, 1, nsxclstr_json['detailed_cluster_status']['overall_status'], style_red)

    online_nodes = len(nsxclstr_json['mgmt_cluster_status']['online_nodes'])

    nsxmgr_url = '/api/v1/cluster'
    nsxmgr_json = session.get(nsx_mgr + str(nsxmgr_url), auth=nsx_auth, verify=session.verify).json()
    nsxmgr_nodes = len(nsxmgr_json["nodes"])
    base = nsxmgr_json["nodes"]
    
    start_row_w = 5
    start_row_x = 6
    start_row_y = 7
    for n in range(online_nodes):
        summary.write(start_row_w, 0, 'NSX-T Manager Appliance FQDN', style_db_left)
        summary.write(start_row_x, 0, 'NSX-T Manager Appliance IP Address', style_db_left)
        summary.write(start_row_y, 0, 'NSX-T Manager Appliance UUID', style_db_left)
        summary.write(start_row_w, 1, base[n]['fqdn'], style_alignleft)
        summary.write(start_row_x, 1, nsxclstr_json['mgmt_cluster_status']['online_nodes'][n]['mgmt_cluster_listen_ip_address'], style_alignleft)
        summary.write(start_row_y, 1, nsxclstr_json['mgmt_cluster_status']['online_nodes'][n]['uuid'], style_alignleft)
        
        start_row_w += 3
        start_row_x += 3
        start_row_y += 3
    
    groups = nsxclstr_json['detailed_cluster_status']['groups']
    
    start_row_z = start_row_y + 1
    
    for n in range(len(groups)):
        summary.write(start_row_x, 0, 'Group ID', style_db_left)
        summary.write(start_row_y, 0, 'Group Type', style_db_left)
        summary.write(start_row_z, 0, 'Group Status', style_db_left)
        summary.write(start_row_x, 1, groups[n]['group_id'], style_alignleft)
        summary.write(start_row_y, 1, groups[n]['group_type'], style_alignleft)
        if groups[n]['group_status'] == 'STABLE':
            summary.write(start_row_z, 1, groups[n]['group_status'], style_green)
        else:
            summary.write(start_row_z, 1, groups[n]['group_status'], style_red)

        mem_row_a = start_row_z + 1
        mem_row_b = mem_row_a + 1
        mem_row_c = mem_row_b + 1
        mem_row_d = mem_row_c + 1

        group_members = groups[n]['members']
        for m in range(len(group_members)):
            summary.write(mem_row_a, 0, 'Member FQDN', style_db_left1)
            summary.write(mem_row_b, 0, 'Member IP address', style_db_left1)
            summary.write(mem_row_c, 0, 'Member UUID', style_db_left1)
            summary.write(mem_row_d, 0, 'Member Status', style_db_left1)
            summary.write(mem_row_a, 1, group_members[m]['member_fqdn'])
            summary.write(mem_row_b, 1, group_members[m]['member_ip'])
            summary.write(mem_row_c, 1, group_members[m]['member_uuid'])
            summary.write(mem_row_d, 1, group_members[m]['member_status'])

            mem_row_a +=4
            mem_row_b +=4
            mem_row_c +=4
            mem_row_d +=4

        start_row_x = mem_row_d - 2
        start_row_y = start_row_x + 1
        start_row_z = start_row_y + 1

    ########### SECTION FOR REPORTING ON INDIVIDUAL MANAGER APPLIANCES ###########

    i = 1
    y = 0

    while i <= nsxmgr_nodes:
        sheet = ls_wkbk.add_sheet('NSX Manager Appliance ' + str(i))
        columnA = sheet.col(0)
        columnA.width = 256 * 30
        columnB = sheet.col(1)
        columnB.width = 256 * 80

        sheet.write(0, 0, 'FQDN', style_db_left)
        sheet.write(1, 0, 'Node ID', style_db_left)
        sheet.write(0, 1, base[y]['fqdn'])
        sheet.write(1, 1, base[y]['node_uuid'])

        entities = len(base[y]['entities'])
        
        row3 = 3
        row4 = 4
        row5 = 5
        
        for n in range(entities):
            sheet.write(row3, 0, 'Entity Type', style_db_left)
            sheet.write(row4, 0, 'IP Address', style_db_left)
            sheet.write(row5, 0, 'Port', style_db_left)

            sheet.write(row3, 1, base[y]['entities'][n]['entity_type'], style_alignleft)
            sheet.write(row4, 1, base[y]['entities'][n]['ip_address'], style_alignleft)
            sheet.write(row5, 1, base[y]['entities'][n]['port'], style_alignleft)

            row3 += 4
            row4 += 4
            row5 += 4

        next_row = row4 - 1
        next_row2 = next_row + 1
        next_row3 = next_row2 + 1
               
        certificates = len(base[y]['certificates'])
        for n in range(certificates):
            sheet.write(next_row, 0, 'Certificate Type', style_db_left)
            sheet.write(next_row2, 0, 'Thumbprint', style_db_left)
            sheet.write(next_row3, 0, 'Certificate', style_db_left)

            sheet.write(next_row, 1, base[y]['certificates'][n]['entity_type'], style_alignleft)
            sheet.write(next_row2, 1, base[y]['certificates'][n]['certificate_sha256_thumbprint'], style_alignleft)
            sheet.write(next_row3, 1, base[y]['certificates'][n]['certificate'], style_alignleft)
            
            next_row += 4
            next_row2 += 4
            next_row3 += 4
        
        y += 1
        i += 1
    
    ls_wkbk.save('NSX Managers.xls')

if __name__ == "__main__":
    main()