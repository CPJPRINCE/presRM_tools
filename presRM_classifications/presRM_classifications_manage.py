import pandas as pd
from pyPreservica import *
import os
import secret
import time
from datetime import datetime

entity = EntityAPI(username=secret.username,password=secret.password, tenant="UARM",server="unilever.preservica.com")
content = ContentAPI(username=secret.username,password=secret.password, tenant="UARM",server="unilever.preservica.com")
retention = RetentionAPI(username=secret.username,password=secret.password, tenant="UARM",server="unilever.preservica.com")

content.search_callback(content.ReportProgressCallBack())

print(entity.token)

startTime = datetime.now()
time.sleep(5)

def list_policies():
    c = 0
    for policy in retention.policies():
        print(policy.name)
        print(policy.reference)
        print(policy.description)
        print(policy.period)
        print(policy.period_unit)
        print(policy.start_date_field)
        print(policy.security_tag)
        print(policy.expiry_action)
        print(policy.restriction)
        print(str(policy.assignable) + "\n")
        c +=1
    print(f"Listed policies: {c}")

#Clearing Policies requires running multiple times... Likely due to list_policies return limits.  
def clear_editable_policies(filter):
    c = 0
    for policy in retention.policies():
        c +=1
        if filter in policy.name:
            try:
                retention.delete_policy(policy.reference)
                print(f"Deleted policy: {policy.reference}, {policy.name}")
            except: print(f'Delete Policy: {policy.reference}, {policy.name} failed, this is fine')
    print(f"Listed policies: {c}")

def create_policies_xl(xl,xltab):
    df = pd.read_excel(xl,xltab)
    dflist = df.values.tolist()
    for rec in dflist:
        print(rec)
        args = dict()
        args['Name'] = str(rec[4])
        args['Description'] = str(rec[5])
        args['SecurityTag'] = "open"
        args['StartDateField'] = "rm.statusdate"
        args['Period'] = str(rec[0])
        args['PeriodUnit'] = str(rec[2])
        args['Restriction'] = "DELETE"
        args['ExpiryAction'] = None
        args['Assignable'] = "true"
        retent = retention.create_policy(**args)
        print(f'Retention created: {retent.reference}')

def list_retention_assignments(ref="%"):
    filters = {"xip.title": "",  "xip.description": "", "xip.document_type": "*",  "xip.parent_ref": "", "xip.security_descriptor": "*",
        "xip.identifier": "", "xip.bitstream_names_r_Preservation": "","xip.retention_policy_assignment_ref":"","xip.retention_policy_assignment_name": "*","rm.statusdate":""}
    search = list(content.search_index_filter_list(query=ref,filter_values=filters))
    print(len(search))
    for hit in search:
        ref = hit['xip.reference']
        asset = entity.asset(ref)
        try:
            assignments = retention.assignments(asset)
            for ass in assignments:
                print(ass.entity_reference, ass.policy_reference, hit['xip.retention_policy_assignment_name'][0], hit['rm.statusdate'],ass.start_date, ass.expired)
        except Exception as e:
            print(f'Failed to get Assignments on: {ref}')
            print(f'Exception: {e}')
            raise SystemError()

def clear_all_retention_assignments(ref="%"):
    filters = {"xip.reference": "", "xip.title": "",  "xip.description": "", "xip.document_type": "*", "xip.parent_hierarchy": "",
            "xip.parent_ref": "", "xip.security_descriptor": "*",
            "xip.identifier": "", "xip.bitstream_names_r_Preservation": "","xip.retention_policy_assignment_ref":"","xip.retention_policy_assignment_name": "*"}
    print(len(list(content.search_index_filter_list(query=ref,filter_values=filters))))
    for hit in list(content.search_index_filter_list(query=ref,filter_values=filters)):
        ref = hit['xip.reference']
        retention_ref = hit['xip.retention_policy_assignment_ref']
        asset = entity.asset(ref)
        try:
            assignments = retention.assignments(asset)
            try:
                print(f'Cleared policy on: {ref}')
                for ass in assignments:
                    retention.remove_assignments(ass)
            except Exception as e:
                print(f'Failed to clear on: {ref} may need to check Secuirty Level')
                print(f'Exception: {e}')
                raise SystemError()                
        except Exception as e:
            print(f'Failed to get Assignments on: {ref}')
            print(f'Exception: {e}')
            raise SystemError()

if __name__ == "__main__":
    xl = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Project Snakeboot\XLoutput\CS10_Extraction_RM-Classifications.xlsx"
    xltab = "PRES_MAP_UK"
    polfilter = "Days"
    #clear_editable_policies(polfilter)
    #list_policies()
    list_retention_assignments("36a67c99-0d29-47ac-ac55-df010927cdb2")
    #create_policies_xl(xl,xltab)
    #clear_all_retention_assignments()
    print(f'Complete, ran for: {datetime.now() - startTime}')