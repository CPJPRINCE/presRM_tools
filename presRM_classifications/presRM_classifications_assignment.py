import pandas as pd
from pyPreservica import *
from .. import secret
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

print(entity.token)

def date_formatting(vardate):
    vardate = str(vardate)
    if not vardate or vardate == "False" or vardate == "?" or vardate == "NaN" or vardate == "nan" or vardate == "NaT" or vardate == "": vardate = ""
    else:
        try:
            vardate = datetime.datetime.strptime(vardate,"%Y-%m-%d %H:%M:%S")
            vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except Exception: 
            try: 
                vardate = datetime.datetime.strptime(vardate,"%Y-%m-%dT%H:%M:%S")
                vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
            except Exception:
                try: 
                    vardate = datetime.datetime.strptime(vardate,"%d-%m-%Y")
                    vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                except Exception:
                    vardate = ""
    return vardate

def parse_text(value):
    if not value or value == "False" or value == "?" or value == "NaN" or value == "nan" or value == "NaT" or value == "": value = "" 
    else: value = str(value).strip()
    return value

def retention_assignment(RETENTION_LIST):
    for RETENT in RETENTION_LIST:
        PRES_REF = RETENT.get('Preservica Reference')
        PRES_RETENT = RETENT.get('Preservica Retention')
        #print(PRES_REF,PRES_RETENT)
        asset_to_assign = entity.asset(PRES_REF)
        class Object(object):
            pass
        rref = Object()
        rref.reference = PRES_RETENT
        try: 
            retention.add_assignments(asset_to_assign,rref)
            print(f'Retention Policy: {PRES_RETENT} Successfully Added to: {PRES_REF}')
        except Exception as e:
            RETENT_ERRORLIST.append(PRES_REF)
            print(e)

if __name__ == '__main__':
    uknl = "UK"
    
    if uknl == "UK":
        nl_flag = False
    elif uknl == "NL":
        nl_flag = True
    else:
        uknl == "UK"
    if nl_flag:  rootname = "Records Management NL"
    else: rootname = "Records Management UK" # Can be overriden with sub-childs for testing...

    xlpath = rf"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Project Snakeboot\XLoutput\CS10Export_{uknl}_Final.xlsx"

    #Lists for Error Checking and for Retention Setting.
    RETENT_ERRORLIST = []
    RETENTION_LIST = []
    #Root Preservica Level, where database will sit. Hard coded, doesn't need to vary - only update on final upload
    PRES_ROOT_FOLDER = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    #PRES_ROOT_FOLDER = "d3d9314b-e1ae-4209-a180-89c07e454456"

    #Data Frame Loading. Sheets needs to be changed on changing from UK to NL. Normally takes ~2 minutes to read UK.
    
    df = pd.read_excel(xlpath,'DATA')
    print(f'Read DataFrame, ran for: {datetime.datetime.now() - startTime}')
    asslist = []
    filters = {"xip.parent_hierarchy":PRES_ROOT_FOLDER,'rm.legacyid':"*","xip.document_type": "IO"}
    content.search_callback(content.ReportProgressCallBack())
    retentionlist = list(content.search_index_filter_list(query="%",filter_values=filters))
    print(f"Hits: {len(retentionlist)}")
    for hit in retentionlist:
        pres_ref = hit['xip.reference']
        idx = df.index[df['id'] == int(hit['rm.legacyid'])].tolist()
        row = df.loc[idx[0]]
        pres_retention = row['classificationpreservica']
        if pres_retention == "Missing" or "Ignore" in pres_retention:
            pass
        else:
            retentdict = {"Preservica Reference": pres_ref,"Preservica Retention": pres_retention}
            retention_assignment([retentdict])
            RETENTION_LIST.append(asslist)
    print(f"Complete! Ran for: {datetime.datetime.now() - startTime}")

    #Lookups the Top-Most Level, to start the Top-Down Loop. - this needs to be set to Records Management UK or Records Management NL.
    #Can also be set to Sub Levels for testing uploads (Rather than changing Tabs.)...
    
    #Retention List Export to Excel w/timestamp.
    df = pd.DataFrame(RETENTION_LIST)
    df.to_excel(f'RETENTION_LIST_{datetime.datetime.now().strftime("%d-%m-%Y")}.xlsx')
    dferr = pd.DataFrame(RETENT_ERRORLIST)
    dferr.to_excel(f'RETENT_ERROR_LIST_{datetime.datetime.now().strftime("%d-%m-%Y")}.xlsx')
