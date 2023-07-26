import pandas as pd
from pyPreservica import *
import secret
from datetime import datetime
import os
import shutil
from lxml import etree
from xml_parsing import XML_Parse,text_formatting,date_formatting
from html import escape

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

def preservica_create_folder(title,description,PRES_PARENT_REF):
    if PRES_PARENT_REF:
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

def metadata_update_box(PRES_REF,xml,loc,NS):
    try:
        new_box = entity.folder(PRES_REF)
        entity.add_identifier(new_box,"code",loc)
        entity.add_metadata(new_box,NS,xml)
    except Exception as e:
        ERROR_LIST.append([PRES_REF,e])
        print(e)

def metadata_update_item(PRES_REF,xml,loc,NS):
    try:
        new_item = entity.asset(PRES_REF)
        entity.add_identifier(new_item,"code",loc)
        entity.add_metadata(new_item,NS,xml)
    except Exception as e:
        ERROR_LIST.append([PRES_REF,e])
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
            RETENT_ERRORLIST.append(PRES_REF,e)

def search_check(search_title,PARENT_REF,desc="",type_flag="Folder",many_flag="create"):
    filters = {"xip.reference":"*","xip.document_type":"*","xip.title":f"{search_title}","xip.parent_ref":f"{PARENT_REF}"}
    search = list(content.search_index_filter_list(query=f"*",filter_values=filters))
    if not len(search):
        SEARCH_REF = PARENT_REF
    elif len(search) > 1:
        print('Too many item\'s have been found; checking for exact match...')
        for s in search:
            if s['xip.title'] == search_title and s['xip.document_type'] == "SO":
                SEARCH_REF = s['xip.reference']
                print(f'Found an exact match: importing to: {SEARCH_REF}')
                break
            else: 
                print('Couldn\'t find an exact match, defaulting to fallback')
                ###### Double Check the Fallback operation...
                if search[0]['xip.document_type'] == "SO": #This is only operating on the first search result... Which might not be true in a multi result.
                    if many_flag == "create":
                        SEARCH_REF = PARENT_REF
                    elif many_flag ==  "first":
                        SEARCH_REF = search[0]['xip.reference']
                    elif many_flag == "donothing":
                        SEARCH_REF = None
                    else: print('Defaulting to First'); SEARCH_REF = search[0]['xip.reference']
                elif search[0]['xip.document_type'] == "IO":
                    SEARCH_REF = None #Sets SEARCH_REF to None if the Asset Already Exists (Multiple times...).
    else:
        #print('Exact Hit... Only 1 Item')
        if search[0]['xip.document_type'] == "SO":
            SEARCH_REF = search[0]['xip.reference']
        elif search[0]['xip.document_type'] == "IO":
            SEARCH_REF = None #Sets SEARCH_REF to None if the Asset Already Exists.
    if SEARCH_REF == PARENT_REF:
        if type_flag == "Folder":
            SEARCH_REF = preservica_create_folder(search_title,desc,SEARCH_REF)
        elif type_flag == "Asset":
            SEARCH_REF = preservica_create_item(search_title,desc,SEARCH_REF)
    elif SEARCH_REF:
        pass
    else:
        print('None Path')
        pass
    return SEARCH_REF

def retention_lookup(assigned_retention,dfclass):
    idx = dfclass.index[dfclass['PolicyName'] == assigned_retention]
    policy = dfclass.loc[idx]
    policy = policy['PolicyRef'].item()
    return policy

def box_item_loop(PARENT_REF):
    if up_flag == "TM":
        print('Running TM Pathway')
        record_dict = dftm.to_dict('records')
        for record in record_dict:
            XML_P = XML_Parse(TMNS)
            xml = XML_P.metadata_parse_TM(record)
            TM_title = str(record.get('Box')) + " - " + str(record.get('Reference')) + " - " + str(record.get('Trademark Name'))
            TM_description = str(record.get('Country')) + " - " + str(record.get('Registration Number')) + " - " + str(record.get('Type of Document'))
            ITEM_PRES_REF = preservica_create_item(TM_title,TM_description,PARENT_REF)
            metadata_update_item(ITEM_PRES_REF,xml,record.get('Location'),TMNS)
            pass
    else:
        box_list = dfbox.to_dict('records')
        for box in box_list:
            type_flag = "Box"
            XML_P = XML_Parse(RMNS)
            box_xml,loc = XML_P.metadata_parse(box_dict=box,item_dict=None,transfer_dict=transfer_dict,type_flag=type_flag)
            BOX_PRES_REF = search_check(box.get('Box Reference'),PARENT_REF,desc=box.get('Description'))
            if not BOX_PRES_REF: pass # If Search Check returns None (Only on 'donothing' tag), 
            else:
                metadata_update_box(BOX_PRES_REF,box_xml,loc,RMNS)
                assigned_retention = box.get('Classification')
                item_list = dfitem.to_dict('records')
                item_search_filters = {"xip.reference":"*","xip.document_type":"IO","xip.title":"*","xip.parent_ref":f"{BOX_PRES_REF}"}
                item_search_dictlist = list(content.search_index_filter_list(query=f"*",filter_values=item_search_filters))
                item_search_list = [x['xip.title'] for x in item_search_dictlist]
                for item in item_list:
                    #for sitem in item_search_list: print(sitem)
                    if any(item.get('Item Reference') in sitem for sitem in item_search_list):
                        print('Item has a match and is already on Preservica... Skipping') 
                        pass
                    else:
                        if str(box.get('Box Reference')) == str(item.get('Box Verify')):
                            if up_flag == "Legal" or up_flag == "Lab" or up_flag == "USCC": type_flag = up_flag
                            else: type_flag = "Item"
                            item_xml, loc = XML_P.metadata_parse(box_dict=box,item_dict=item,transfer_dict=transfer_dict,type_flag=type_flag)
                            ITEM_PRES_REF = preservica_create_item(item.get('Item Reference'),item.get('Description'),BOX_PRES_REF)
                            #ITEM_PRES_REF = search_check(itemref,BOX_PRES_REF,desc=itemdesc,type_flag="Asset")
                            #if not ITEM_PRES_REF: pass #Skips Items that are already uploaded onto System.
                            #else: 
                            metadata_update_item(ITEM_PRES_REF,item_xml,loc,RMNS)
                            policy = retention_lookup(assigned_retention,dfclass)
                            RETENTION_LIST.append({"Preservica Reference":ITEM_PRES_REF,"Preservica Retention":policy})
                        else: 
                            print('Item does not have a matching Box Verify Reference... Skipping.')
                            pass

def make_structure(folder):
    if os.path.exists(folder): pass
    else: os.makedirs()


def test_search():
    test_reference = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    test_title = "Legal"

    filters = {"xip.reference":"*","xip.document_type":"*","xip.title":f"{test_title}","xip.parent_ref":f"{test_reference}"}
    search = list(content.search_index_filter_list(query=f"*",filter_values=filters))
    print(search)
    for s in search:
        print(s['xip.title'])
        if s['xip.title'] == test_title: print('Yes')
    raise SystemExit

if __name__ == '__main__':

    upload_folder = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Preservica Spreadsheet Uploads\Uploads - PLACE COMPLETED TEMPLATE IN HERE"
    #upload_folder = "/Users/archives/Unilever/UARM Teamsite - Records Management/Preservica Spreadsheet Uploads/Uploads - PLACE COMPLETED TEMPLATE IN HERE"
    error_folder = os.path.join(upload_folder,"errors")    
    #test_search()

    RMNS = "http://rm.unilever.co.uk/schema"
    TMNS = "http://rm.unilever.co.uk/trademarks/v1"

    ERROR_LIST = []
    RETENT_ERRORLIST = []
    
    PRES_ROOT_FOLDER = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    TMAGR_ROOT = "b4a34349-f5d8-4096-bd44-8a2a30e7482c"
    TMCERT_ROOT = "07c847df-f708-408b-b7d4-0ed4ae48de31"
    LAB_ROOT = "e7bb6aa3-e19c-4d6b-8e0b-d62c471b176e"
    UKI_ROOT = "fabf56bf-c0ce-4d43-9a7e-6699048df575"
    UNI_ROOT = "7c770368-ac4f-4701-9dad-6bea58585a97"
    USCC_ROOT = ""

    temp_folder = os.path.join(upload_folder,"temp")
    make_structure(temp_folder)
    comp_folder = os.path.join(upload_folder,"complete")
    make_structure(comp_folder)

    for file in os.listdir(upload_folder):
        RETENTION_LIST = [] 
        file_path = os.path.join(upload_folder,file)
        if os.path.isfile(file_path) and file_path.endswith('.xlsx'):
            print(f'Processing: {file_path}')
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            temp_path = os.path.join(temp_folder,timestamp + "_" + file)
            comp_path = os.path.join(comp_folder,timestamp + "_" + file)           
            shutil.move(file_path,temp_path)
            if os.path.exists(temp_path): print(f'Successfully Moved to Temp: {temp_path}')
            else: print('Something has gone wrong moving to Temp Folder... Raising Exit'); raise SystemExit()
            xl = pd.ExcelFile(temp_path)
            dftop = xl.parse('Transfer')
            transfer_dict = dftop.to_dict('records')[0]
            if "TM" in xl.sheet_names:
                dftm = xl.parse('TM')
                up_flag = "TM"
            else: 
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Item')
                dfclass = xl.parse('VAL-CLASS')
            xl.close()
            if "Legal" in xl.sheet_names:
                up_flag = "Legal"
                legal_dept_search = transfer_dict.get('Top-Level')
                if legal_dept_search == "UKI": TOP_REF = UKI_ROOT
                elif legal_dept_search == "UNI": TOP_REF = UNI_ROOT
                else: print('Top-Level needs to be either UNI or UKI; raising System Exit'); raise SystemExit()
                box_item_loop(TOP_REF)
            #Lab Notebooks will default to 
            elif "Lab" in xl.sheet_names:
                up_flag = "Lab"
                TOP_REF = LAB_ROOT
                box_item_loop(TOP_REF)
            elif "TM" in xl.sheet_names:
                if transfer_dict.get('Database') == "Certificates": TOP_REF = TMCERT_ROOT
                elif transfer_dict.get('Database') == "Agreements": TOP_REF = TMAGR_ROOT
                else: print('Something Fishy is going on, and I don\'t like it one bit...'); raise SystemExit()
                box_item_loop(TOP_REF)
            else:
                up_flag = False
                top_search = transfer_dict.get('Top-Level Folder')
                TOP_REF = search_check(top_search,PRES_ROOT_FOLDER,desc=top_search)
                dept_search = transfer_dict.get('Departmental Folder')
                DEPT_REF = search_check(dept_search,TOP_REF,desc=dept_search)
                date_title = datetime.strftime((transfer_dict.get('Transfer Date')),"%d/%m/%Y")
                date_desc = date_title
                DATE_REF = search_check(date_title, DEPT_REF, desc=date_title)
                box_item_loop(DATE_REF)
            shutil.move(temp_path,comp_path)
            print(f'Successfully moved Temp_File: {temp_path}; to: {comp_path}')
            print(f'Processed Structure Upload, ran for: {datetime.now() - startTime}')
            print(f'Sleeping 10 Secs, before running Retention Assignment...')
            time.sleep(10)
            retention_assignment(RETENTION_LIST)
        else: 'Ignoring Non-Excel file'
    print(f"Complete!, ran for: {datetime.now() - startTime}")
    if ERROR_LIST:
        dferr = pd.DataFrame(ERROR_LIST,columns=['Pres Ref','Error Message'])
        with pd.ExcelWriter(os.path.join(error_folder,f"Error List_{timestamp}",".xlsx"),engine='auto',mode='w') as xlWriter:
            dferr.to_excel(xlWriter,'Errors',index=False)
    if RETENT_ERRORLIST:
        dferr = pd.DataFrame(RETENT_ERRORLIST,columns=['Pres Ref','Error Message'])
        with pd.ExcelWriter(os.path.join(error_folder,f"Retention_Error List_{timestamp}",".xlsx"),engine='auto',mode='w') as xlWriter:
            dferr.to_excel(xlWriter,'Errors',index=False)