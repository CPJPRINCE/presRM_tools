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
    if not value or value == "False" or value == "?" or str(value) == "NaN" or str(value) == "nan" or str(value) == "NaT" or value == "": value = "" 
    else: value = str(value).strip()
    return value

def metadata_parse_review(row,prow=pd.Series(dtype="float64")):
    if row.empty:
        row = prow
    if not row.empty:
        row = row.replace({'&':'&amp;','\<': '&lt;','\>':'&gt;',"\'":'&apos;','\"':'&quot;'},regex=True)
        row = row.fillna(str(''))
        datereview = date_formatting(row['reviewsdate'])
        reviewsentby = parse_text(row['reviewssentby'])
        decision = parse_text(row['reviewsaction'])
        authorisedby = parse_text(row['reviewsauthorisedby'])
        notes1 = parse_text(row['reviewsdecisionnotes'])
        notes2 = parse_text(row['reviewsextendednotes'])
        nextdate = date_formatting(row['reviewsnewdate'])
        datedecision = date_formatting(row['reviewssigneddate'])
        if not datedecision: datedecision = date_formatting(row['reviewsextendedon'])
        if not notes2 and notes1:
            notes = f"Decision Notes: {notes1}"
        elif not notes1 and notes2:
            notes = f"Extension Notes : {notes2}"
        elif notes1 and notes2:
            notes = f"Decision Notes: {notes1} \n Extended Notes: {notes2}"
        else: notes = ""
        xml_parse_rev = f'<rmreview:rmreview xmlns:rmreview="{RMNS_REV}" xmlns="{RMNS_REV}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' \
            f"<rmreview:datereview>{datereview}</rmreview:datereview>" \
            f"<rmreview:reviewnumber></rmreview:reviewnumber>" \
            f"<rmreview:datereviewsent></rmreview:datereviewsent>" \
            f"<rmreview:reviewsentby>{reviewsentby}</rmreview:reviewsentby>" \
            f"<rmreview:decision>{decision}</rmreview:decision>" \
            f"<rmreview:datedecision>{datedecision}</rmreview:datedecision>" \
            f"<rmreview:datenextreview>{nextdate}</rmreview:datenextreview>" \
            f"<rmreview:authorisedby>{authorisedby}</rmreview:authorisedby>" \
            f"<rmreview:formerrsi></rmreview:formerrsi>" \
            f"<rmreview:reviewnotes>{notes}</rmreview:reviewnotes>" \
            f"</rmreview:rmreview>"            
    else: xml_parse_rev = xml_parse_rev
    try: xml_test = etree.fromstring(xml_parse_rev)
    except Exception as e: 
        print(f"XML Exception: {e}")
        print(xml_parse_rev)    
    return xml_parse_rev

def metadata_parse_removal(row,prow=pd.Series(dtype="float64")):
    if row.empty:
        row = prow
    if not row.empty:
        row = row.replace({'&':'&amp;','\<': '&lt;','\>':'&gt;',"\'":'&apos;','\"':'&quot;'},regex=True)
        row = row.fillna(str(''))
        datereview = date_formatting(row['removalsdate'])
        deptname = parse_text(row['removalsdepartment'])
        authorisedby = parse_text(row['removalsauthorisedby'])
        actionedby = parse_text(row['removalsactionedby'])
        removalnotes = parse_text(row['removalsnotes'])
        xml_parse_rem = f'<rmremove:rmremove xmlns:rmremove="{RMNS_REM}" xmlns="{RMNS_REM}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' \
        f"<rmremove:dateremove>{datereview}</rmremove:dateremove>" \
        f"<rmremove:departmentname>{deptname}</rmremove:departmentname>" \
        f"<rmremove:authorisedby>{authorisedby}</rmremove:authorisedby>" \
        f"<rmremove:actionedby>{actionedby}</rmremove:actionedby>" \
        f"<rmremove:removalnotes>{removalnotes}</rmremove:removalnotes>" \
        "</rmremove:rmremove>"
    else: xml_parse_rem = xml_parse_rem
    try: xml_test = etree.fromstring(xml_parse_rem)
    except Exception as e: 
        print(f"XML Exception: {e}")
        print(xml_parse_rem)
    return xml_parse_rem


if __name__ == '__main__':

    RMNS_REM = "http://rm.unilever.co.uk/removals/v1"
    RMNS_REV = "http://rm.unilever.co.uk/review/v1"

    xlpath = rf"C:\Users\Chris.Prince\OneDrive - Unilever\Documents\CS10Export_UK_Missing Reviews_Test.xlsx"
    
    df = pd.read_excel(xlpath,sheet_name='data',index_col='index')
    df = df.reset_index()
    dffold = pd.read_excel(xlpath,sheet_name='datawfolders',index_col='index')
    dfold = dffold.reset_index()
    pres_list = df['Preservica_Ref'].values.tolist()
    for pres_ref in pres_list:
        if not str(row['reviewsdate']) == "nan" or str(row['reviewsaction']) == "nan":
            idx = df.index[df['Preservica_Ref'] == pres_ref]
            row = df.loc[idx[0]]
            try: 
                pidx = dffold.index(dffold['id'] == row['parentid'])
                prow = dffold.loc[pidx[0]]
            except:
                prow=pd.Series(dtype="float64")
            try: 
                ent = entity.entity(EntityType.ASSET,pres_ref)
            except Exception as e:
                if "404" in str(e):
                    ent = entity.entity(EntityType.FOLDER,pres_ref)
                else: print(e)
            xml_str = metadata_parse_review(row,prow)
            entity.delete_metadata(ent,RMNS_REV)
            entity.add_metadata(ent,RMNS_REV,xml_str)
    if not str(row['removalsdate']) == "nan":
        try: 
            ent = entity.entity(EntityType.ASSET,pres_ref)
        except Exception as e:
            if "404" in str(e):
                ent = entity.entity(EntityType.FOLDER,pres_ref)
            else: print(e)
        xml_str = metadata_parse_removal(row,prow)
        entity.delete_metadata(ent,RMNS_REM)
        entity.add_metadata(ent,RMNS_REM,xml_str)