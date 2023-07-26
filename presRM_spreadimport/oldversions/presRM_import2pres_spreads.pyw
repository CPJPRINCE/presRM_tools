import pandas as pd
from pyPreservica import *
import secret
import re
from datetime import datetime,timezone
import os
import shutil
from lxml import etree
import presRM_label_generation as labelgen

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


def parse_text(value):
    value = str(value)
    if value == "False" or value == "NaN" or value == "nan" or value == "NaT" or value == "": value = ""
    else: value = str(value).strip()
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
    for PRES_REF,RETENTION_REF in RETENTION_LIST:
        #print(PRES_REF,RETENTION_REF)
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


def metadata_parse(row,prow=pd.Series(dtype="float64")):
    #Escaping Characters.
    row = row.replace({'&':'&amp;','\<': '&lt;','\>':'&gt;',"\'":'&apos;','\"':'&quot;'},regex=True)
    row = row.fillna(str(''))
    if not prow.empty:
        prow = prow.replace({'&':'&amp;','<': '&lt;','>':'&gt;',"'":'&apos;','"':'&quot;'},regex=True)
        prow = prow.fillna(str(''))

    #Object Type information... A legacy part of CS10 imported into Preservica. 
    obj = parse_text(row['objtype'])
    if obj == "Physical Item": obj = "Item"
    elif obj == "Physical Item Container": obj = "Box"
    elif obj == "Folder": obj = "Folder"
    elif obj == "Document": obj = "Electronic Record"
    else: obj = obj

    #Location information - sets to MISSING if... missing
    loc = parse_text(row['boxlocation'])
    if loc == str(''): 
        loc = "MISSING"

    statdate = date_formatting(row['statusdate'])
    # Contact Info
    droclient = parse_text(row['client'])
    alternativecontact = parse_text(row['altcontact'])
    
    #Media type pulled through to determine Item Type
    mediatype = parse_text(row['mediatype'])
    numberofitems = parse_text(row['numberofitems'])
    coverdates = parse_text(row['coverdates'])
    format = parse_text(row['format'])

    # Parsing Paret Row - Casacades partial info down if prow variable is provided.
    # Only intended for Physical Items.
    if not prow.empty:
        loc = parse_text(prow['boxlocation'])
        defaultloc = parse_text(prow['homelocation'])
        locatortype = ""
        area = parse_text(prow['area'])
        weight = ""
    else:
        defaultloc = parse_text(row['homelocation'])
        locatortype = parse_text(row['locatortype'])
        area = parse_text(row['area'])
        loc = parse_text(row['boxlocation'])
        weight = parse_text(row['weight'])

    #Area... May need to vary
    if area == "MAIN": area = "UK RMS MAIN"
    elif area == "CAGE": area = "UK RMS CAGE"
    else: area = area

    #Legal Item Info
    legaldoc = parse_text(row['legdocttpye'])
    copy = parse_text(row['legcopy'])
    party1 = parse_text(row['partyname1'])
    party2 = parse_text(row['partyname2'])
    party3 = parse_text(row['partyname3'])
    party4 = parse_text(row['partyname4'])
    dateofagreement = date_formatting(row["legdate"])
    dateoftermination = date_formatting(row["legtermdate"])
    amendments = parse_text(row["legamends"])
    value = parse_text(row["legvalue"])
    currency = parse_text(row["legcurrency"])
    datesubmitted = parse_text(row['datesubmitted'])
    agreementobjectives = parse_text(row['agreeobj'])

    #USCC Info
    usccagreementtype = parse_text(row['uscctype'])
    agreementowner = parse_text(row['usccowner'])
    destructiondate = date_formatting(row['usccdestruct'])
    agreementend = date_formatting(row['usccend'])
    docnum = parse_text(row['usccref'])
    agreementstart = date_formatting(row['usccstart'])
    vendordesc = parse_text(row['usccvenddesc'])
    vendorcode = parse_text(row['usccvendname'])
    upccode = parse_text(row['usccuuc'])
    usccdatesubmitted = parse_text(row['datesubmitted'])

    #Lab Info
    labtype = parse_text(row['uscctype'])
    labnotebook = parse_text(row['labnotebook'])
    author = parse_text(row['labauth'])
    department = parse_text(row['labdept'])
    site = parse_text(row['labsite'])
    dateofissue = date_formatting(row['labdateissue'])
    dateofcompletion = date_formatting(row['datecomp'])
    
    #LegalDept defining UNI or UKI
    legaldept = ""
    
    #Also defining itemtype
    if "Legal" in mediatype or "legal" in mediatype:
        itemtype = "Legal Agreement"
        if "UKI" in mediatype: legaldept = "UKI"
        elif "UNI" in mediatype: legaldept = "UNI"
        else: legaldept = ""
    elif "USCC" in mediatype: itemtype = "USCC Agreement Transmittal Form"
    elif "Lab" in mediatype: itemtype = "Lab Notebook"
    elif mediatype == "UK Item": itemtype = "Item"
    elif mediatype == "UK Series": itemtype = "Series"
    elif mediatype == "UK Store Resource": itemtype = "Placeholder"
    elif mediatype == "UK RMbox": itemtype = "RM Box"
    elif mediatype == "UK Archive box": itemtype = "Archive Box"
    else: itemtype = "Other"
    
    #XML String Construction, Split into parts, mainly for visual.

    xml_parse_pt1 = f'<rm:rm xmlns:rm="{RMNS}" xmlns="{RMNS}">'\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statdate}</rm:statusdate>" \
                f"<rm:coverdate>{coverdates}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{obj}</rm:objtype>" \
                "</rm:recordinfo>" \
                "<rm:clientinfo>" \
                f"<rm:client>{droclient}</rm:client>"\
                f"<rm:alternatecontact>{alternativecontact}</rm:alternatecontact>" \
                "</rm:clientinfo>" \
                "<rm:boxinfo>" \
                f"<rm:weight>{weight}</rm:weight>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                f"<rm:boxtype>{locatortype}</rm:boxtype>" \
                "</rm:boxinfo>" \
                "<rm:iteminfo>" \
                f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                f"<rm:format>{format}</rm:format>" \
                f"<rm:numberofitems>{numberofitems}</rm:numberofitems>"

    xml_item_add = "<rm:legalitem>" \
                f"<rm:legaldept>{legaldept}</rm:legaldept>" \
                f"<rm:legaldoc>{legaldoc}</rm:legaldoc>" \
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
    
    #Transfer Information
    transferdate = date_formatting(row['transferdate'])
    transferaccepted = parse_text(row['transferacceptedby'])
    transfernotes = parse_text(row['transfernotes'])
    transfercharge = parse_text(row['transfercharge'])
    transferredby = parse_text(row['transfertransferedby'])

    #Legacy Info
    legacycreatedby = parse_text(row['createdby'])
    legacycreateddate = date_formatting(row['datecreated'])
    legacyid = parse_text(row['id'])
    legacypid = parse_text(int(row['parentid']))
    alternativereference = parse_text(row['altreference'])

    xml_parse_pt2 = "<rm:transferinfo>" \
                f"<rm:transferdate>{transferdate}</rm:transferdate>" \
                f"<rm:transferaccepted>{transferaccepted}</rm:transferaccepted>" \
                f"<rm:transfernotes>{transfernotes}</rm:transfernotes>" \
                f"<rm:transfercharge>{transfercharge}</rm:transfercharge>" \
                f"<rm:transferredby>{transferredby}</rm:transferredby>" \
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

    return xml_parse,loc

def metadata_parse_box(metadata_list):
    objtype = "Box"
    boxref = parse_text(metadata_list[0])
    boxdesc = parse_text(metadata_list[1])
    client = parse_text(metadata_list[2])
    altcont = parse_text(metadata_list[3])
    cvrdate = parse_text(metadata_list[4])
    weight = parse_text(metadata_list[5])
    loc = parse_text(metadata_list[6])
    boxtype = parse_text(metadata_list[8])
    statusdate = date_formatting(metadata_list[9])
    defaultloc = parse_text(metadata_list[10])
    area = parse_text(metadata_list[11])
    xml_parse = f"<rm:rm xmlns:rm='{RMNS}' xmlns='{RMNS}'>"\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                f"<rm:coverdate>{cvrdate}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{objtype}</rm:objtype>" \
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
                "</rm:rm>"
    xml_test = etree.fromstring(xml_parse)

    return xml_parse,boxref,boxdesc,loc

def metadata_parse_item(metadata_list,box_metadata):
    objtype = "Item"
    itemtype = parse_text(metadata_list[0])
    itemref = parse_text(metadata_list[2])
    itemdesc = parse_text(metadata_list[3])
    format = parse_text(metadata_list[4])
    cvrdate = parse_text(metadata_list[5])
    numofitems = parse_text(metadata_list[6])
    #Box Data
    loc = parse_text(box_metadata[6])
    altcont = parse_text(box_metadata[4])
    client = parse_text(box_metadata[2])
    statusdate = date_formatting(box_metadata[9])
    area = parse_text(box_metadata[11])
    defaultloc = parse_text(box_metadata[10])

    xml_parse = f"<rm:rm xmlns:rm='{RMNS}' xmlns='{RMNS}'>"\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                f"<rm:coverdate>{cvrdate}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{objtype}</rm:objtype>" \
                "</rm:recordinfo>" \
                "<rm:clientinfo>" \
                f"<rm:client>{client}</rm:client>"\
                f"<rm:alternatecontact>{altcont}</rm:alternatecontact>" \
                "</rm:clientinfo>" \
                "<rm:boxinfo>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                "</rm:boxinfo>" \
                "<rm:iteminfo>" \
                f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                f"<rm:format>{format}</rm:format>" \
                f"<rm:numberofitems>{numofitems}</rm:numberofitems>" \
                "</rm:iteminfo>" \
                "</rm:rm>"
    xml_test = etree.fromstring(xml_parse)

    return xml_parse,itemref,itemdesc

def metadata_parse_legal(metadata_list,box_metadata):
    objtype = "Item"
    itemtype = "Legal Agreement"
    itemref = parse_text(metadata_list[1])
    itemdesc = parse_text(metadata_list[2])
    cvrdate = parse_text(metadata_list[3])
    agreetype = parse_text(metadata_list[4])
    agreeother = parse_text(metadata_list[5])
    copy = parse_text(metadata_list[6])
    part1 = parse_text(metadata_list[7])
    part2 = parse_text(metadata_list[8])
    part3 = parse_text(metadata_list[9])
    part4 = parse_text(metadata_list[10])
    dateagree = date_formatting(metadata_list[11])
    dateterm = date_formatting(metadata_list[12])
    amend = parse_text(metadata_list[13])
    value = parse_text(metadata_list[14])
    currency = parse_text(metadata_list[15])
    notes = parse_text(metadata_list[16])
    datesubmit = date_formatting(metadata_list[17])
    agreeobj = parse_text(metadata_list[18])
    #Box Data
    loc = parse_text(box_metadata[6])
    altcont = parse_text(box_metadata[4])
    client = parse_text(box_metadata[2])
    statusdate = date_formatting(box_metadata[9])
    area = parse_text(box_metadata[11])
    defaultloc = parse_text(box_metadata[10])

    xml_parse = f"<rm:rm xmlns:rm='{RMNS}' xmlns='{RMNS}'>"\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                f"<rm:coverdate>{cvrdate}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{objtype}</rm:objtype>" \
                f"<rm:notes>{notes}</rm:notes>" \
                "</rm:recordinfo>" \
                "<rm:clientinfo>" \
                f"<rm:client>{client}</rm:client>"\
                f"<rm:alternatecontact>{altcont}</rm:alternatecontact>" \
                "</rm:clientinfo>" \
                "<rm:boxinfo>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                "</rm:boxinfo>" \
                "<rm:iteminfo>" \
                f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                f"<rm:legalitem>" \
                f"<rm:legaldept>{top_folder}</rm:legaldept>" \
                f"<rm:legaldoc>{agreetype}</rm:legaldoc>" \
                f"<rm:legaldocother>{agreeother}</rm:legaldocother>" \
                f"<rm:copy>{copy}</rm:copy>" \
                f"<rm:partyname1>{part1}</rm:partyname1>" \
                f"<rm:partyname2>{part2}</rm:partyname2>" \
                f"<rm:partyname3>{part3}</rm:partyname3>" \
                f"<rm:partyname4>{part4}</rm:partyname4>" \
                f"<rm:dateofagreement>{dateagree}</rm:dateofagreement>" \
                f"<rm:dateoftermination>{dateterm}</rm:dateoftermination>" \
                f"<rm:amendments>{amend}</rm:amendments>" \
                f"<rm:value>{value}</rm:value>" \
                f"<rm:currency>{currency}</rm:currency>" \
                f"<rm:datesubmitted>{datesubmit}</rm:datesubmitted>" \
                f"<rm:agreementobjectives>{agreeobj}</rm:agreementobjectives>" \
                f"</rm:legalitem>" \
                "</rm:iteminfo>" \
                "</rm:rm>"
    xml_test = etree.fromstring(xml_parse)

    return xml_parse,itemref,itemdesc

def metadata_parse_lab(metadata_list,box_metadata):
    objtype = "Item"
    itemtype = parse_text(metadata_list[0])
    itemref = parse_text(metadata_list[1])
    itemdesc = parse_text(metadata_list[2])
    cvrdate = parse_text(metadata_list[3])
    notebook = parse_text(metadata_list[4])
    author = parse_text(metadata_list[5])
    department = parse_text(metadata_list[6])
    site = parse_text(metadata_list[7])
    dateofissue = parse_text(metadata_list[8])
    dateofcompletion = parse_text(metadata_list[9])
    notes = parse_text(metadata_list[10])
    #Box Data
    loc = parse_text(box_metadata[6])
    altcont = parse_text(box_metadata[4])
    client = parse_text(box_metadata[2])
    statusdate = date_formatting(box_metadata[9])
    area = parse_text(box_metadata[11])
    defaultloc = parse_text(box_metadata[10])

    xml_parse = f"<rm:rm xmlns:rm='{RMNS}' xmlns='{RMNS}'>"\
                "<rm:recordinfo>" \
                f"<rm:statusdate>{statusdate}</rm:statusdate>" \
                f"<rm:coverdate>{cvrdate}</rm:coverdate>" \
                f"<rm:disposition>false</rm:disposition>" \
                f"<rm:objtype>{objtype}</rm:objtype>" \
                f"<rm:notes>{notes}</rm:notes>" \
                "</rm:recordinfo>" \
                "<rm:clientinfo>" \
                f"<rm:client>{client}</rm:client>"\
                f"<rm:alternatecontact>{altcont}</rm:alternatecontact>" \
                "</rm:clientinfo>" \
                "<rm:boxinfo>" \
                f"<rm:area>{area}</rm:area>" \
                f"<rm:boxlocation>{loc}</rm:boxlocation>" \
                f"<rm:defaultlocation>{defaultloc}</rm:defaultlocation>" \
                "</rm:boxinfo>" \
                "<rm:iteminfo>" \
                f"<rm:itemtype>{itemtype}</rm:itemtype>" \
                f"<rm:labitem>" \
                f"<rm:notebook>{notebook}</rm:notebook>" \
                f"<rm:author>{author}</rm:author>" \
                f"<rm:department>{department}</rm:department>" \
                f"<rm:site>{site}</rm:site>" \
                f"<rm:dateofissue>{dateofissue}</rm:dateofissue>" \
                f"<rm:dateofcompletion>{dateofcompletion}</rm:dateofcompletion>" \
                f"</rm:labitem>" \
                "</rm:iteminfo>" \
                "</rm:rm>"
    xml_test = etree.fromstring(xml_parse)

    return xml_parse,itemref,itemdesc

def processing(REF):
    box_list = dfbox.values.tolist()
    for box in box_list:
        box_meta,boxref,boxdesc,loc = metadata_parse_box(box)
        BOX_PRES_REF = preservica_create_box(boxref,boxdesc,REF)
        metadata_update_box(BOX_PRES_REF,box_meta,loc)
        assigned_retention = box[7]
        # if check for multiple boxes?...
        item_list = dfitem.values.tolist()
        for item in item_list:
                if str(boxref) == str(item[1]):
                    if type_flag == "Legal": item_meta, itemref, itemdesc = metadata_parse_legal(item,box)  
                    elif type_flag == "Lab": item_meta,  itemref, itemdesc = metadata_parse_lab(item,box)              
                    else: item_meta, itemref, itemdesc = metadata_parse_item(item,box)
                    ITEM_PRES_REF = preservica_create_item(itemref,itemdesc,BOX_PRES_REF)
                    metadata_update_item(ITEM_PRES_REF,item_meta,loc)
                    policy = retention_lookup(assigned_retention,dfclass)
                    RETENTION_LIST.append([ITEM_PRES_REF,policy])
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
    PRES_ROOT_FOLDER = "e36534dc-b0b9-4cc2-88f5-8d34f670f9c2"
    
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
            if "Legal" in xl.sheet_names:
                type_flag = "Legal"
                dftop = xl.parse('Transfer')
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Legal')
                dfclass = xl.parse('VAL-CLASS')
                xl.close()
                top_folder = dftop['Top-Level'].item()
                if top_folder == "UKI": TOP_REF = "c1082d21-0b6a-4be5-9ffa-f821ddf056d3" ##### ADD PRESERVICA FOLDER REF FOR UKI
                elif top_folder == "UNI": TOP_REF = "4fc6f4bc-063b-4a97-9c04-ba49b188ff26" ##### ADD PRESERVICA FOLDER REF FOR UNI
                else: print('Top-Level needs to be either UNI or UKI; raising System Exit'); raise SystemExit()
                processing(TOP_REF)
            elif "Lab" in xl.sheet_names:
                type_flag = "Lab"
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Lab')
                dfclass = xl.parse('VAL-CLASS')
                xl.close()
                TOP_REF = "e795710b-f9ff-4e6a-a46d-d9e0ac9dd599"  ##### ADD PRESERVICA FOLDER REF FOR LAB
                processing(TOP_REF)
            else:
                type_flag = False
                dftop = xl.parse('Transfer')
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Item')
                dfclass = xl.parse('VAL-CLASS')
                xl.close()
                top_search = dftop['Top-Level Folder'].item()
                TOP_REF = search_check(top_search,PRES_ROOT_FOLDER)
                dept_search = dftop['Departmental Folder'].item()
                DEPT_REF = search_check(dept_search,TOP_REF)
                # if DEPT_REF == "CREATEFOLDER":
                #     pass
                date_title = datetime.strftime((dftop['Transfer Date'].item()),"%d/%m/%Y")
                date_desc = date_title
                DATE_REF = search_check(date_title, DEPT_REF)
                if DATE_REF == DEPT_REF:
                    DATE_REF = preservica_create_folder(date_title,date_desc,DEPT_REF)
                    processing(DATE_REF)
                else:
                    processing(DATE_REF)
            shutil.move(temp_path,comp_path)
            print(f'Processed Structure Upload, ran for: {datetime.now() - startTime}')
            print(f'Running Retention Assignment...')
            retention_assignment(RETENTION_LIST)

    print(f"Complete!, ran for: {datetime.now() - startTime}")