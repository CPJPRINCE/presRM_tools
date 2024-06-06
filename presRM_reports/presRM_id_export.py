import pandas as pd
from pyPreservica import *
import secret
from datetime import datetime
from lxml import etree
import pandas as pd

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

def path_parent_return(ref,io):
    if io == "IO":
        x = entity.asset(ref)
    else: 
        x = entity.folder(ref)
    fpath = x.title        
    f = entity.folder(x.parent)
    parent_name = f.title
    while f.parent is not None:
        f = entity.folder(f.parent)
        fpath = f.title + ":::" + fpath
        fpath = fpath.rsplit(":::",maxsplit=1)[0]
        fpath = fpath.split(":::",maxsplit=2)[-1]
    return fpath,parent_name

if __name__ == '__main__':

    #Root Preservica Level, where database will sit.
    PRES_ROOT_FOLDER = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    #PRES_ROOT_FOLDER = "3cea8f86-1f4e-43dc-81c3-cfa81d5370f0"

    PATH_FLAG = False
    #Data Frame Loading. Sheets needs to be changed on changing from UK to NL. Normally takes ~2 minutes to read UK.
    
    filters = {"xip.parent_hierarchy":PRES_ROOT_FOLDER,'rm.legacyid':"*","xip.reference": "*","xip.parent_hierachy": "*","xip.title": "*","xip.description":"","xip.document_type":"*","xip.retention_policy_assignment_name": ""}
    dict_list = []
    content.search_callback(content.ReportProgressCallBack())
    report_search = content.search_index_filter_list(page_size=1000,query="%", filter_values=filters)

    for hit in report_search:
        ref = hit['xip.reference']
        if PATH_FLAG:
            full_path,parent_title = path_parent_return(ref,hit['xip.document_type'])
            dict_ent = {'Reference': ref, 'Path': full_path, 'Parent':parent_title,'Title': hit['xip.title'],'Description':hit['xip.description'],'LegacyID':hit['rm.legacyid'],'Policy Name': hit['xip.retention_policy_assignment_name']}
        else: dict_ent = {'Reference': ref, 'Title': hit['xip.title'],'Description':hit['xip.description'],'LegacyID':hit['rm.legacyid'],'Policy Name': hit.get('xip.retention_policy_assignment_name')}
        dict_list.append(dict_ent)
    df = pd.DataFrame(dict_list)
    df.to_excel('presRM_UK_id.xlsx')
    print("Finish Time: " + str(datetime.now() - startTime))
