import pandas as pd
from pyPreservica import *
import secret
import re
from datetime import datetime
import xlwings
import os
import shutil

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

upload_folder = r"C:\Users\Chris\Unilever\UARM Teamsite - Archive\Digital Preservation\github\cs10_extraction\Uploads"

xlpath = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Project Snakeboot\XLoutput\CS10_Extraction_RM-UK_Processing.xlsx"

RMNS = "http://rm.unilever.co.uk/schema"

def metadata_parse(row,prow=pd.Series(dtype="float64")):
    row = row.replace("&","&amp;")
    row = row.fillna(str(''))
    if not prow.empty:
        prow = prow.replace("&","&amp;")
        prow = prow.fillna(str(''))
    obj = row['Attribute:objname']
    loc = row['podata.Attribute:boxlocator']
    if loc == str(''): 
        loc = "MISSING"
        #print('MISSING!')
    statdate = str(row['StatusDate'])
    if statdate == "False" or statdate == "NaN" or statdate == "nan" or statdate == "": statdate = datetime.now()
    else: statdate = datetime.strptime(statdate,"%d-%m-%Y")
    recdate = str(row['RecordDate'])
    if recdate == "False" or recdate == "NaN" or recdate == "nan" or recdate == "":
        recdate = datetime.now()
    else: recdate = datetime.strptime(recdate,"%d-%m-%Y")

    droclient = row['podata.Attribute:client'] ### Why have these not pulled through?
    if not str(row['client.telephone']) == "nan" or str(row['client.telephone']) == "NaN" or str(row['client.telephone']) == "": dronumber = str(row['client.telephone'])
    else: dronumber = None
    if not str(row['client.email']) == "nan" or str(row['client.email']) == "NaN" or str(row['client.email']) == "": droemail = str(row['client.email'])
    else: droemail = None
    if not str(row['client.address']) == "nan" or str(row['client.address']) == "NaN" or str(row['client.address']) == "": droaddress = str(row['client.address'])
    else: droaddress = None

    currentloc = row['podata.Attribute:CurrentLocation']
    defaultloc = row['podata.Attribute:defaultlocation']
    mediatype = row['podata.Attribute:MediaTypeName']
    alternativecontact = row['Alternative contact']
    coverdates = row['Covering dates']
    format = row['Format']

    legacycreatedby = row['Attribute:createdbyname']
    legacycreateddate = row['Attribute:created']
    legacyid = row['Attribute:id']
    legacypid = row['Attribute:parentid']

    if not prow.empty:
        locatortype = prow['podata.Attribute:locatortype']
        area = prow['podata.Attribute:area']
        loc = prow['podata.Attribute:boxlocator']
        weight = prow['Weight (kg)']
    else:
        locatortype = row['podata.Attribute:locatortype']
        area = row['podata.Attribute:area']
        loc = row['podata.Attribute:boxlocator']
        weight = row['Weight (kg)']


    # recdate, statdate, disposition needs calculating
    # drodept,droeml,droadr need lookup from other spreadsheet
    # import objtype? loc? mediatype?

    
    xml_parse = f"<rm:rm xmlns:rm='{RMNS}' xmlns='{RMNS}'>"\
                "<rm:recordinfo>" \
                f"<rm:recordmanager></rm:recordmanager>" \
                f"<rm:recorddate>{recdate}</rm:recorddate>" \
                f"<rm:statusdate>{statdate}</rm:statusdate>" \
                f"<rm:coverdate>{coverdates}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{obj}</rm:objtype>" \
                "</rm:recordinfo>" \
                "<rm:departmentinfo>" \
                f"<rm:dro>{droclient}</rm:dro>"\
                f"<rm:dronumber>{dronumber}</rm:dronumber>" \
                f"<rm:droemail>{droemail}</rm:droemail>" \
                f"<rm:droaddress>{droaddress}</rm:droaddress>" \
                f"<rm:dronotes></rm:dronotes>" \
                f"<rm:alternatecontact>{alternativecontact}</rm:alternatecontact>" \
                "</rm:departmentinfo>" \
                "<rm:formercontactinfo>" \
                f"<rm:formerdro></rm:formerdro>" \
                f"<rm:formerhead></rm:formerhead>" \
                f"<rm:formerdepartmentname></rm:formerdepartmentname>" \
                f"<rm:departmenthistorynotes></rm:departmenthistorynotes>" \
                "</rm:formercontactinfo>" \
                "<rm:boxinfo>" \
                f"<rm:weight>{weight}</rm:weight>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:format>{format}</rm:format>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                f"<rm:currentlocation>{currentloc}</rm:currentlocation>" \
                f"<rm:boxtype>{locatortype}</rm:boxtype>" \
                f"<rm:mediatype>{mediatype}</rm:mediatype>" \
                "</rm:boxinfo>" \
                "<rm:transferinfo>" \
                f"<rm:transferdate></rm:transferdate>" \
                f"<rm:transferaccepted></rm:transferaccepted>" \
                f"<rm:transfernotes></rm:transfernotes>" \
                f"<rm:transfercharge></rm:transfercharge>" \
                f"<rm:transferredby></rm:transferredby>" \
                "</rm:transferinfo>" \
                "</rm:rm>"

    return xml_parse,loc

def preservica_create_folder(title,description,PRES_PARENT_REF):
    new_folder = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
    return new_folder.reference

def preservica_create_box(title,description,PRES_PARENT_REF):
    new_box = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
    return new_box.reference

def preservica_create_item(title,description,PRES_PARENT_REF):
    parent = entity.folder(PRES_PARENT_REF)
    new_item = entity.add_physical_asset(title,description,parent,"closed")
    return new_item.reference

def preservica_upload_electronic(title,description,PRES_PARENT_REF):
    pass

def metadata_update_box(PRES_REF,metadata):
    try:
        new_box = entity.folder(PRES_REF)
        entity.add_identifier(new_box,"code",metadata[1])
        entity.add_metadata(new_box,RMNS,metadata[0])
    except Exception as e:
        ErrorList.append(PRES_REF)
        print(e)

def metadata_update_item(PRES_REF,metadata):
    try:
        new_item = entity.asset(PRES_REF)
        entity.add_identifier(new_item,"code",metadata[1])
        entity.add_metadata(new_item,RMNS,metadata[0])
    except Exception as e:
        ErrorList.append(PRES_REF)
        print(e)


def retention_assignment(RETENTION_LIST):
    for PRES_REF,RETENTION_REF in RETENTION_LIST:
        asset_to_assign = entity.asset(PRES_REF)
        class Object(object):
            pass
        rref = Object()
        rref.reference = RETENTION_REF
        try: 
            retention.add_assignments(asset_to_assign,rref)
        except Exception as e:
            print(e)

def search_for_folder(parent_title,PARENT_PRES):
    filters = {"xip.reference":"*","xip.title":f"{parent_title}","xip.parent_ref":f"{PARENT_PRES}"}
    print(parent_title)
    search = list(content.search_index_filter_list(query=f"{parent_title}",filter_values=filters))
    if not len(search):
        print('No hits')
        dest_folder = "CREATEFOLDER"
    elif len(search) > 1:
        print('Too many')
        dest_folder = "TOO MANY HITS"
    else:
        print('Just right')
        dest_folder = search[0]['xip.reference']
    return dest_folder

def search_check(search_name,PRES_REF):
    search_check = search_for_folder(search_name,PRES_REF)
    if search_check == "CREATEFOLDER":
        SEARCH_PRES_REF = PRES_REF
    elif search_check == "TOO MANY HITS":
        print('TOO MANY HITS... SOMETHING\'S AWRY')
        print('Will default to creating a new folder')
        SEARCH_PRES_REF = PRES_REF
        pass
    else: SEARCH_PRES_REF = search_check
    return SEARCH_PRES_REF

def topdown_loop(id,PARENT_PRES_REF,c=0):
    idx = df.index[df['Attribute:parentid'] == id].tolist()
    for i in idx:
        print('Processing: ' + str(c),end="\r")
        c +=1
        if not c % 500:
            print(f"\nRunning for: {datetime.now() - startTime}\n")
        row = df.loc[i]
        row_name = str(row['name.Element:Text'])
        row_description = str(row['description.Element:Text'])
        if row_description == "nan":
            row_description = row_name
        row_id = row['Attribute:id']
        row_pid = row['Attribute:parentid']
        row_objname = str(row['Attribute:objname'])
        row_retentionassignment = str(row['Preservica.PolicyRef'])
        if row_objname == "Folder":
            ### print('I\'m a Folder!') #Debug
            ### Enable and Shift this level, to Eliminate 'Date' Level
            # if re.match("[0-9]{2}/[0-9]{2}/[0-9]{4}",row_name):
            #     print('Skipping Date Folder')
            #     SUB_PRES_REF = PARENT_PRES_REF
            # else:
            ###Testing for Search use IE Template Upload...
            search_check()
            SUB_PRES_REF = str(preservica_create_folder(row_name,row_description,PARENT_PRES_REF))
            ### print(F"Created Folder: {row_name}, in Preservica at: {SUB_PRES_REF}") #Debug
            topdown_loop(row_id,SUB_PRES_REF,c)
        elif row_objname == "Physical Item Container":
            ### print('I\'m a Box!') #Debug
            PRES_BOX_REF = str(preservica_create_box(row_name,row_description,PARENT_PRES_REF))        
            metadata = metadata_parse(row)
            metadata_update_box(PRES_BOX_REF,metadata)
            ### print(F"Created Box: {row_name}, in Preservica at: {PRES_BOX_REF}") #Debug
            topdown_loop(row_id,PRES_BOX_REF,c)
        elif row_objname =="Physical Item":
            ### print('I\'m an Item!') #Debug
            PRES_ITEM_REF = preservica_create_item(row_name,row_description,PARENT_PRES_REF)
            pidx = df.index[df['Attribute:id'] == row_pid]
            prow = df.loc[pidx[0]]
            metadata = metadata_parse(row,prow)
            ### print(F"Created Item: {row_name}, in Preservica at: {PRES_ITEM_REF}") #Debug
            metadata_update_item(PRES_ITEM_REF,metadata)
            if not row_retentionassignment == "NaN" or not row_retentionassignment == "nan":
                RETENTION_LIST.append([PRES_ITEM_REF,row_retentionassignment])
            else: print('Retention is not available...')
        elif row_objname == "Document":
            ### print('I\'m a Electronic Record!') #Debug
            ### print('Doing Nothing...') #Debug
            pass
        else: print('I\m nothing! Or something that shouldn\'t be here...')
        #Loop was out of place here... would iterate 'flatly' over all items, causing break of tree, only needs to reiterate for Folders/Containers... 
        #topdown_loop(row_id,ROOT_PRES_REF,level)

if __name__ == '__main__':
    ErrorList = []
    RETENTION_LIST = [] 
    PRES_ROOT_FOLDER = "e36534dc-b0b9-4cc2-88f5-8d34f670f9c2"
    
    for f in os.listdir(upload_folder):
        fpath = os.path.join(upload_folder,f)
        print(fpath)
        if os.path.isfile(fpath) and fpath.endswith('.xlsx'):
           print(fpath)
           dftop = pd.read_excel(fpath,'TOP')
           dfbox = pd.read_excel(fpath,'BOX')
           dfitem = pd.read_excel(fpath,'ITEM')
           top_search = dftop['Top-Level Folder'].item()
           TOP_REF = search_check(top_search,PRES_ROOT_FOLDER)
           dept_search = dftop['Departmental Folder'].item()
           DEPT_REF = search_check(dept_search,TOP_REF)
           if not DEPT_REF == "CREATEFOLDER":
               pass
           date_title = datetime.strftime((dftop['Date'].item()),"%d/%m/%Y")
           date_desc = date_title
           DATE_REF = search_check(date_title, DEPT_REF)
           if DATE_REF == "CREATEFOLDER":
               DATE_REF = preservica_create_folder(date_title,date_desc,DEPT_REF)
               time.sleep(1)
               boxtitle = dfbox['Title'].item()
               boxdesc = dfbox['Description'].item()
               boxref = dfbox['Box Reference'].item()
               BOX_REF = preservica_create_box(boxtitle,boxdesc,DATE_REF)
               item_list = dfitem.values.tolist()
               for f in item_list
           else: print(DATE_REF)


    print(f'Read DataFrame, ran for: {datetime.now() - startTime}')
    
    #topdown_loop(start_id, PRES_ROOT_FOLDER)
    #df = pd.DataFrame(RETENTION_LIST,columns=['Reference','Retention'])
    #df.to_excel(f'Retention_list_{datetime.now().strftime("%d-%M-%Y")}.xlsx')
    #retention_assignment(RETENTION_LIST)
    #dferr = pd.DataFrame(ErrorList,columns=["Reference"])
    #dferr.to_excel(f'Error_list_{datetime.now().strftime("%d-%M-%Y")}.xlsx')
    
    print('Complete!')
    print(f"Ran for: {datetime.now() - startTime}")