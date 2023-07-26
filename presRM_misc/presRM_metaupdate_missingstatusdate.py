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
                    vardate = datetime.datetime.strptime(vardate,"%d/%m/%Y")
                    vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                    vardate = ""
    return vardate

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


    RMNS = "http://rm.unilever.co.uk/schema"
    xlpath = rf"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Project Snakeboot\XLoutput\pres_missing_statusdates.xlsx"

    df = pd.read_excel(xlpath,sheet_name='Sheet1')
    pres_list = df[['PresReference','statusdate']].values.tolist()
    for pres_ref,statdate in pres_list:
        try: 
            ent = entity.entity(EntityType.ASSET,pres_ref)
        except Exception as e:
            if "404" in str(e):
                ent = entity.entity(EntityType.FOLDER,pres_ref)
            else: print(e)
        xml_string = entity.metadata_for_entity(ent,RMNS)
        xml_doc = etree.fromstring(xml_string)
        statd = xml_doc.find('.//{http://rm.unilever.co.uk/schema}statusdate')
        statdate = date_formatting(statdate)
        statd.text = str(statdate)
        xml_out = etree.tostring(xml_doc,encoding="UTF-8").decode("utf-8")
        entity.update_metadata(ent,RMNS,xml_out)