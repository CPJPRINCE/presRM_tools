import pandas as pd
from pyPreservica import *
import secret
import re
from datetime import datetime,timezone
import os
import shutil
from lxml import etree

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

def preservica_create_folder(title,description,PRES_PARENT_REF):
    new_folder = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
    print(f'Created Folder at: {title}, Reference {new_folder.reference}')
    return new_folder.reference

def preservica_create_box(title,description,PRES_PARENT_REF):
    new_box = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
    print(f'Created Box at: {title}, Reference {new_box.reference}')
    return new_box.reference

def preservica_create_item(title,description,PRES_PARENT_REF):
    parent = entity.folder(PRES_PARENT_REF)
    new_item = entity.add_physical_asset(title,description,parent,"closed")
    print(f'Created Item at: {title}, Reference {new_item.reference}')
    return new_item.reference

def preservica_upload_electronic(title,description,PRES_PARENT_REF):
    pass


def parse_text(value):
    value = str(value)
    if value == "False" or value == "NaN" or value == "nan" or value == "NaT" or value == "": value = ""
    else: 
        value = str(value).strip()
        value = value.replace("&","&amp;")
        value = value.replace(">","&gt;")
        value = value.replace("<","&lt;")
        value = value.replace("'","&apos;")
        value = value.replace("\"","&quot;")
        #value = value.replace({'&':'&amp;','\<': '&lt;','\>':'&gt;',"\'":'&apos;','\"':'&quot;'},regex=True)
    return value

def date_formatting(vardate):
    vardate = str(vardate)
    if vardate == "False" or vardate == "NaN" or vardate == "nan" or vardate == "NaT" or vardate == "": vardate = datetime.now().strftime("%Y-%m-%dT%H:%M:%S.000Z")
    else: vardate = datetime.strptime(vardate,"%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return vardate

def metadata_update_box(PRES_REF,xml,loc):
    try:
        new_box = entity.folder(PRES_REF)
        entity.add_identifier(new_box,"code",loc)
        entity.add_metadata(new_box,RMNS,xml)
    except Exception as e:
        ERROR_LIST.append(PRES_REF)
        print(e)

def metadata_update_item(PRES_REF,xml,loc):
    try:
        new_item = entity.asset(PRES_REF)
        entity.add_identifier(new_item,"code",loc)
        entity.add_metadata(new_item,RMNS,xml)
    except Exception as e:
        ERROR_LIST.append(PRES_REF)
        print(e)

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

def search_for_folder(parent_title,PARENT_PRES):
    filters = {"xip.reference":"*","xip.title":f"{parent_title}","xip.parent_ref":f"{PARENT_PRES}"}
    print(parent_title)
    search = list(content.search_index_filter_list(query=f"*",filter_values=filters))
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

def retention_lookup(assigned_retention,dfclass):
    idx = dfclass.index[dfclass['PolicyName'] == assigned_retention]
    policy = dfclass.loc[idx]
    policy = policy['PolicyRef'].item()
    return policy

def metadata_parse(box_dict,item_dict=None,type_flag=None,transfer_dict=None):
    #Escaping Characters.
    transdate = ""
    transaccpt = ""
    transby = ""
    transcharge = ""
    transnotes = ""

    objtype = ""
    title = ""
    desc = ""

    loc = ""
    weight = ""
    altcont = ""
    client = ""
    statusdate = ""
    boxtype = ""
    area = ""
    defaultloc = ""
    coverdates = ""
    notes = ""
    
    itemtype = ""
    numberofitems = ""
    coverdates = ""
    format = ""
    
    labtype = ""
    labnotebook = ""
    author = ""
    department = ""
    site = ""
    dateofissue = ""
    dateofcompletion  = ""

    usccagreementtype = ""
    legaldept = ""
    legaldoc = ""
    legaldocother = ""
    copy = ""
    party1 = ""
    party2 = ""
    party3 = ""
    party4 = ""
    dateofagreement = ""
    dateoftermination  = ""
    amendments = ""
    value = ""
    currency = ""
    datesubmitted = ""
    agreementobjectives = ""
        
    usccagreementtype = ""
    agreementowner = ""
    destructiondate = ""
    agreementend = ""
    docnum = ""
    agreementstart = ""
    vendordesc = ""
    vendorcode = ""
    upccode = ""
    usccdatesubmitted = ""

    #Legacy Info - Will always be set to Null.

    legacycreatedby = ""
    legacycreateddate = ""
    legacyid = ""
    legacypid = ""
    alternativereference = ""

    if type_flag == "Box":
        objtype = "Box"
        title = parse_text(box_dict.get('Box Reference'))
        desc = parse_text(box_dict.get('Description'))        
        loc = parse_text(box_dict.get('Location'))
        weight = parse_text(box_dict.get('Weight'))
        altcont = parse_text(box_dict.get('Alternative Contact'))
        client = parse_text(box_dict.get('Client'))
        statusdate = date_formatting(box_dict.get('Status Date'))
        boxtype = parse_text(box_dict.get('Box Type'))
        area = parse_text(box_dict.get('Area'))
        defaultloc = parse_text(box_dict.get('Home Location'))
        coverdates = parse_text(box_dict.get('Covering Dates'))
        notes = parse_text(box_dict.get('Notes'))

        #Transfer_Dict only Get's fed into Box Level 
        transdate = date_formatting(transfer_dict.get('Transfer Date'))
        transaccpt = parse_text(transfer_dict.get('Accepted By'))
        transby = parse_text(transfer_dict.get('Transferred By'))
        transcharge = parse_text(transfer_dict.get('Chargeable'))
        transnotes = parse_text(transfer_dict.get('Transfer Notes'))       

    elif type_flag == "Item":
        objtype = "Item"
        #Box Information Recurses down to Item Level
        loc = parse_text(box_dict.get('Location'))
        altcont = parse_text(box_dict.get('Alternative Contact'))
        client = parse_text(box_dict.get('Client'))
        statusdate = date_formatting(box_dict.get('Status Date'))
        area = parse_text(box_dict.get('Area'))
        defaultloc = parse_text(box_dict.get('Home Location'))
        #Item Level Info
        title = parse_text(item_dict.get('Item Reference'))
        desc = parse_text(item_dict.get('Description'))        
        itemtype = parse_text(item_dict.get('Item Type'))
        numberofitems = parse_text(item_dict.get('Number of Items'))
        coverdates = parse_text(item_dict.get('Covering Dates'))
        format = parse_text(item_dict.get('Format'))
        notes = parse_text(item_dict.get('Notes'))

    elif type_flag == "Lab":
        objtype = "Item"
        #Box Information Recurses down to Item Level
        loc = parse_text(box_dict.get('Location'))
        altcont = parse_text(box_dict.get('Alternative Contact'))
        client = parse_text(box_dict.get('Client'))
        statusdate = date_formatting(box_dict.get('Status Date'))
        area = parse_text(box_dict.get('Area'))
        defaultloc = parse_text(box_dict.get('Home Location'))
        #Lab - Base
        title = parse_text(box_dict.get('Title'))
        desc = parse_text(box_dict.get('Description'))        
        coverdates = parse_text(item_dict.get('Covering Dates'))
        itemtype = parse_text(item_dict.get('Item Type'))
        notes = parse_text(item_dict.get('Notes'))
        #Lab - Unq
        labtype = parse_text(item_dict.get('Lab Type'))
        labnotebook = parse_text(item_dict.get('Lab Notebook Number'))
        author = parse_text(item_dict.get('Author'))
        department = parse_text(item_dict.get('Department'))
        site = parse_text(item_dict.get('Site'))
        dateofissue = date_formatting(item_dict.get('Date of Issue'))
        dateofcompletion = date_formatting(item_dict.get('Date of Completion'))
    elif type_flag == "Legal":
        objtype = "Item"
        #Box Information Recurses down to Item Level
        loc = parse_text(box_dict.get('Location'))
        altcont = parse_text(box_dict.get('Alternative Contact'))
        client = parse_text(box_dict.get('Client'))
        statusdate = date_formatting(box_dict.get('Status Date'))
        area = parse_text(box_dict.get('Area'))
        defaultloc = parse_text(box_dict.get('Home Location'))
        #Legal - Base
        itemtype = parse_text(item_dict.get('Item Type'))
        title = parse_text(item_dict.get('Agreement Reference'))
        desc = parse_text(item_dict.get('Description'))        
        coverdates = parse_text(item_dict.get('Covering Dates'))
        notes = parse_text(item_dict.get('Notes'))
        #Legal - Unq
        legaldept = parse_text(transfer_dict.get('Top-Level')) # Makes Special Call on Transfer_Dict
        legaldoc = parse_text(item_dict.get('Agreement Type'))
        legaldocother = parse_text(item_dict.get('If Other, please state'))
        copy = parse_text(item_dict.get('Copy'))
        party1 = parse_text(item_dict.get('Party Name 1'))
        party2 = parse_text(item_dict.get('Party Name 2'))
        party3 = parse_text(item_dict.get('Party Name 3'))
        party4 = parse_text(item_dict.get('Party Name 4'))
        dateofagreement = date_formatting(item_dict.get('Date of Agreement'))
        dateoftermination = date_formatting(item_dict.get('Date of Termination'))
        amendments = parse_text(item_dict.get('Amendments'))
        value = parse_text(item_dict.get('Value'))
        currency = parse_text(item_dict.get('Currency'))
        if legaldept == "UKI":
            datesubmitted = date_formatting(item_dict.get('Date Submitted (UKI)'))
            agreementobjectives = parse_text(item_dict.get('Agreement Objectives (UKI)'))
        else: datesubmitted = ""; agreementobjectives = ""

    # elif type_flag == "USCC": ### USCC Needs Template Defining...
    #     #USCC Info
    #     itemtype = "USCC"
    #     usccagreementtype = parse_text(row['uscctype'])
    #     agreementowner = parse_text(row['usccowner'])
    #     destructiondate = date_formatting(row['usccdestruct'])
    #     agreementend = date_formatting(row['usccend'])
    #     docnum = parse_text(row['usccref'])
    #     agreementstart = date_formatting(row['usccstart'])
    #     vendordesc = parse_text(row['usccvenddesc'])
    #     vendorcode = parse_text(row['usccvendname'])
    #     upccode = parse_text(row['usccuuc'])
    #     usccdatesubmitted = parse_text(row['datesubmitted'])

    elif type_flag == "TM":
        objtype = "Item"
    else: print('Invalid Type'); pass

    xml_parse_pt1 = f'<rm:rm xmlns:rm="{RMNS}" xmlns="{RMNS}">'\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                f"<rm:coverdate>{coverdates}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{objtype}</rm:objtype>" \
                f"<rm:recordnotes>{notes}</rm:recordnotes>" \
                "</rm:recordinfo>" \
                "<rm:clientinfo>" \
                f"<rm:client>{client}</rm:client>"\
                f"<rm:alternatecontact>{altcont}</rm:alternatecontact>" \
                "</rm:clientinfo>" \
                "<rm:boxinfo>" \
                f"<rm:weight>{weight}</rm:weight>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                f"<rm:boxtype>{boxtype}</rm:boxtype>" \
                "</rm:boxinfo>" \
                "<rm:iteminfo>" \
                f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                f"<rm:format>{format}</rm:format>" \
                f"<rm:numberofitems>{numberofitems}</rm:numberofitems>"

    xml_item_add = "<rm:legalitem>" \
                f"<rm:legaldept>{legaldept}</rm:legaldept>" \
                f"<rm:legaldoc>{legaldoc}</rm:legaldoc>" \
                f"<rm:legaldocother>{legaldocother}</rm:legaldocother>" \
                f"<rm:copy>{copy}</rm:copy>" \
                f"<rm:partyname1>{party1}</rm:partyname1>" \
                f"<rm:partyname2>{party2}</rm:partyname2>" \
                f"<rm:partyname3>{party3}</rm:partyname3>" \
                f"<rm:partyname4>{party4}</rm:partyname4>" \
                f"<rm:dateofagreement>{dateofagreement}</rm:dateofagreement>" \
                f"<rm:dateoftermination>{dateoftermination}</rm:dateoftermination>" \
                f"<rm:amendments>{amendments}</rm:amendments>" \
                f"<rm:value>{value}</rm:value>" \
                f"<rm:currency>{currency}</rm:currency>" \
                f"<rm:datesubmitted>{datesubmitted}</rm:datesubmitted>" \
                f"<rm:agreementobjectives>{agreementobjectives}</rm:agreementobjectives>" \
                "</rm:legalitem>" \
                "<rm:usccitem>" \
                f"<rm:agreementtype>{usccagreementtype}</rm:agreementtype>" \
                f"<rm:upccode>{upccode}</rm:upccode>" \
                f"<rm:vendorcode>{vendorcode}</rm:vendorcode>" \
                f"<rm:vendordesc>{vendordesc}</rm:vendordesc>" \
                f"<rm:agreementowner>{agreementowner}</rm:agreementowner>" \
                f"<rm:agreementstart>{agreementstart}</rm:agreementstart>" \
                f"<rm:agreementend>{agreementend}</rm:agreementend>" \
                f"<rm:docnumber>{docnum}</rm:docnumber>" \
                f"<rm:destructiondate>{usccdatesubmitted}</rm:destructiondate>" \
                f"<rm:destructiondate>{destructiondate}</rm:destructiondate>" \
                "</rm:usccitem>" \
                "<rm:labitem>" \
                f"<rm:agreementtype>{labtype}</rm:agreementtype>" \
                f"<rm:notebook>{labnotebook}</rm:notebook>" \
                f"<rm:author>{author}</rm:author>" \
                f"<rm:department>{department}</rm:department>" \
                f"<rm:site>{site}</rm:site>" \
                f"<rm:dateofissue>{dateofissue}</rm:dateofissue>" \
                f"<rm:dateofcompletion>{dateofcompletion}</rm:dateofcompletion>" \
                "</rm:labitem>" \
                "</rm:iteminfo>"

    xml_parse_pt2 = "<rm:transferinfo>" \
                f"<rm:transferdate>{transdate}</rm:transferdate>" \
                f"<rm:transferaccepted>{transaccpt}</rm:transferaccepted>" \
                f"<rm:transfernotes>{transnotes}</rm:transfernotes>" \
                f"<rm:transfercharge>{transcharge}</rm:transfercharge>" \
                f"<rm:transferredby>{transby}</rm:transferredby>" \
                "</rm:transferinfo>" \
                "<rm:legacyinfo>" \
                f"<rm:createdby>{legacycreatedby}</rm:createdby>" \
                f"<rm:createddate>{legacycreateddate}</rm:createddate>" \
                f"<rm:legacyid>{legacyid}</rm:legacyid>" \
                f"<rm:legacypid>{legacypid}</rm:legacypid>" \
                f"<rm:alternativereference>{alternativereference}</rm:alternativereference>" \
                "</rm:legacyinfo>" \
                "</rm:rm>"

    xml_parse = xml_parse_pt1 + xml_item_add + xml_parse_pt2
    
    try: xml_test = etree.fromstring(xml_parse)
    except Exception as e: 
        print(e)
        print(xml_parse)
    return xml_parse,title,desc,loc

def processing(REF):
    box_list = dfbox.to_dict('records')
    for box in box_list:
        type_flag = "Box"
        box_meta,boxref,boxdesc,loc = metadata_parse(box_dict=box,item_dict=None,transfer_dict=transfer_dict,type_flag=type_flag)
        BOX_PRES_REF = preservica_create_box(boxref,boxdesc,REF)
        metadata_update_box(BOX_PRES_REF,box_meta,loc)
        assigned_retention = box.get('Classification')
        # if check for multiple boxes?...
        item_list = dfitem.to_dict('records')
        for item in item_list:
                if str(boxref) == str(item.get('Box Verify')):
                    if up_flag == "Legal" or up_flag == "Lab" or up_flag == "USCC": type_flag = up_flag
                    else: type_flag = "Item"
                    item_meta, itemref, itemdesc, loc = metadata_parse(box_dict=box,item_dict=item,transfer_dict=transfer_dict,type_flag=type_flag)  
                    ITEM_PRES_REF = preservica_create_item(itemref,itemdesc,BOX_PRES_REF)
                    metadata_update_item(ITEM_PRES_REF,item_meta,loc)
                    policy = retention_lookup(assigned_retention,dfclass)
                    RETENTION_LIST.append({"Preservica Reference":ITEM_PRES_REF,"Preservica Retention":policy})
                else:
                    #Skips Items that aren't a match to the boxref
                    pass

if __name__ == '__main__':

    upload_folder = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Preservica Spreadsheet Uploads\Uploads - PLACE COMPLETED TEMPLATE IN HERE"
    #upload_folder = "/Users/archives/Unilever/UARM Teamsite - Records Management/Preservica Spreadsheet Uploads/Uploads - PLACE COMPLETED TEMPLATE IN HERE"
    
    temp_folder = os.path.join(upload_folder,"temp")
    if os.path.exists(temp_folder): pass
    else: os.makedirs(temp_folder)
    comp_folder = os.path.join(upload_folder,"complete")
    if os.path.exists(comp_folder): pass
    else: os.makedirs(comp_folder)

    RMNS = "http://rm.unilever.co.uk/schema"
    ERROR_LIST = []
    RETENTION_LIST = [] 
    RETENT_ERRORLIST = []
    PRES_ROOT_FOLDER = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    
    for file in os.listdir(upload_folder):
        file_path = os.path.join(upload_folder,file)
        if os.path.isfile(file_path) and file_path.endswith('.xlsx'):
            print(f'Processing: {file_path}')
            date_now = datetime.now().strftime("%Y%m%d%H%M%S")
            temp_path = os.path.join(temp_folder,date_now + "_" + file)
            comp_path = os.path.join(comp_folder,date_now + "_" + file)           
            shutil.move(file_path,temp_path)
            if os.path.exists(temp_path): print('Successfully Moved to Temp Folder')
            else: print('Something has gone wrong moving to Temp Folder... Raising Exit'); raise SystemExit()
            xl = pd.ExcelFile(temp_path)
            dftop = xl.parse('Transfer')
            transfer_dict = dftop.to_dict('records')[0]
            dfbox = xl.parse('Box')
            dfitem = xl.parse('Item')
            dfclass = xl.parse('VAL-CLASS')
            xl.close()
            if "Legal" in xl.sheet_names:
                up_flag = "Legal"
                legal_dept_search = transfer_dict.get('Top-Level')
                if legal_dept_search == "UKI": TOP_REF = "fabf56bf-c0ce-4d43-9a7e-6699048df575"
                elif legal_dept_search == "UNI": TOP_REF = "7c770368-ac4f-4701-9dad-6bea58585a97"
                else: print('Top-Level needs to be either UNI or UKI; raising System Exit'); raise SystemExit()
                processing(TOP_REF)
            elif "Lab" in xl.sheet_names:
                up_flag = "Lab"
                TOP_REF = "e7bb6aa3-e19c-4d6b-8e0b-d62c471b176e"
                processing(TOP_REF)
            else:
                up_flag = False
                top_search = transfer_dict.get('Top-Level Folder')
                TOP_REF = search_check(top_search,PRES_ROOT_FOLDER)
                dept_search = transfer_dict.get('Departmental Folder')
                DEPT_REF = search_check(dept_search,TOP_REF)
                # if DEPT_REF == "CREATEFOLDER":
                #     pass
                date_title = datetime.strftime((transfer_dict.get('Transfer Date')),"%d/%m/%Y")
                date_desc = date_title
                DATE_REF = search_check(date_title, DEPT_REF)
                if DATE_REF == DEPT_REF:
                    DATE_REF = preservica_create_folder(date_title,date_desc,DEPT_REF)
                    processing(DATE_REF)
                else:
                    processing(DATE_REF)
            shutil.move(temp_path,comp_path)
            print(f'Processed Structure Upload, ran for: {datetime.now() - startTime}')
            print(f'Sleeping 10 Secs, Running Retention Assignment...')
            time.sleep(10)
            retention_assignment(RETENTION_LIST)

    print(f"Complete!, ran for: {datetime.now() - startTime}")