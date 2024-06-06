from pyPreservica import *
from datetime import datetime 
import csv
import secret

content = ContentAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")
#"rm.statusdate":"1900-01-01 - 3000-01-01"
filters = {"xip.title":"","xip.document_type":"IO","xip.retention_policy_assignment_name":"", "xip.parent_hierarchy":"bf614a5d-10db-4ff8-addb-2c913b3149eb"}

list_omit = ["aa1d8512-766b-400a-bdbe-f97aeb487418",
             "c3eb7e2b-9615-4fb3-a51a-4d3eb8d67ce2",
             "eecd6de2-f039-4d8f-85bc-8944089fbe2a",
             "62ccc74c-2d2a-4769-bf62-aa97d1960d49"]

content.search_callback(content.ReportProgressCallBack())

now = datetime.now()
search = list(content.search_index_filter_list(query="%",filter_values=filters,page_size=1000))

list_missingretentions = []
set_keys = set()
for hit in search:
    if hit.get('retention_policy_assignment_name') is None and not any(hit.get('xip.reference') in omit for omit in list_omit):
        list_missingretentions.append(hit)
        set_keys.add(list(hit.keys()))

if len(list_missingretentions) == 0: print('No Retention Policies are missing')
else:
    keys = list_missingretentions[0].keys()
    with open("missing retentionpolicies report_" \
            + datetime.strftime(now,"%Y%m%d") \
            + ".csv",'w',encoding="utf-8",newline='') as output_file:
        csvwrite = csv.DictWriter(output_file,set_keys)
        csvwrite.writeheader()
        csvwrite.writerows(list_missingretentions)
