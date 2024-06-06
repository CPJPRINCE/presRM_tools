from pyPreservica import *
from datetime import datetime 
import csv
import secret

content = ContentAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")
#"rm.statusdate":"1900-01-01 - 3000-01-01"
filters = {"xip.title":"","xip.retention_policy_assignment_name":"*","rm.statusdate":"", "xip.parent_hierarchy":"bf614a5d-10db-4ff8-addb-2c913b3149eb"}

content.search_callback(content.ReportProgressCallBack())

now = datetime.now()
search = list(content.search_index_filter_list(query="%",filter_values=filters,page_size=1000))

list_missingstatdate = [hit for hit in search if hit.get('rm.statusdate') is None]
if len(list_missingstatdate) == 0: print('No status Dates are missing')
else:
    keys = list_missingstatdate[0].keys()
    with open("missing statusdate report_" + datetime.strfltime(now,"%Y%m%d") + ".csv",'w',encoding="utf-8",newline='') as output_file:
        csvwrite = csv.DictWriter(output_file,keys)
        csvwrite.writeheader()
        csvwrite.writerows(list_missingstatdate)
