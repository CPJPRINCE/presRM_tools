import pandas as pd
from pyPreservica import *
import secret
import datetime
from lxml import etree

startTime = datetime.datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

if __name__ == '__main__':
        
    filters = {"xip.parent_hierarchy":"bf614a5d-10db-4ff8-addb-2c913b3149eb","xip.reference": "","xip.description":"_x000D_"}
    content.search_callback(content.ReportProgressCallBack())
    search = list(content.search_index_filter_list(query="%", filter_values=filters))
    for hit in search:
        folder = entity.folder(hit.get('xip.reference'))
        desc = hit.get('xip.description')
        desc.replace('_x000D_\n',"")
        print(desc)
