import pandas as pd
from pyPreservica import *
import secret
from datetime import datetime
import os
import shutil
from lxml import etree
from xml_parsing import XML_Parse,text_formatting,date_formatting
#from html import escape
from re import escape

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

def preservica_create_folder(title,description,PRES_PARENT_REF):
    title = text_formatting(title)
    description = text_formatting(description)
    if PRES_PARENT_REF:
        new_folder = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
        print(f'Created Folder at: {title}, Reference {new_folder.reference}')
    return new_folder.reference

def preservica_create_box(title,description,PRES_PARENT_REF):
    title = text_formatting(title)
    description = text_formatting(description)
    new_box = entity.create_folder(title,description,"closed",PRES_PARENT_REF)
    print(f'Created Box at: {title}, Reference {new_box.reference}')
    return new_box.reference

def preservica_create_item(title,description,PRES_PARENT_REF):
    title = text_formatting(title)
    description = text_formatting(description)
    parent = entity.folder(PRES_PARENT_REF)
    new_item = entity.add_physical_asset(title,description,parent,"closed")
    print(f'Created Item at: {title}, Reference {new_item.reference}')
    return new_item.reference

def preservica_upload_electronic(title,description,PRES_PARENT_REF):
    pass

def metadata_update_box(PRES_REF,xml,loc,NS):
    try:
        new_box = entity.folder(PRES_REF)
        idents = entity.identifiers_for_entity(new_box)
        for ident in idents:
            if "code" in ident: break    
        else: entity.add_identifier(new_box,"code",loc)
        existing_xml = entity.metadata_for_entity(new_box,RMNS)
        if existing_xml:
            xml = xml_merge(xml_a=etree.fromstring(existing_xml),xml_b=etree.fromstring(xml))
            entity.update_metadata(new_box,NS,xml.decode('utf-8'))
        else:
            entity.add_metadata(new_box,NS,xml)
    except Exception as e:
        ERROR_LIST.append([PRES_REF,e])
        print(e)
        raise SystemError()

def metadata_update_item(PRES_REF,xml,loc,NS):
    try:
        new_item = entity.asset(PRES_REF)
        idents = entity.identifiers_for_entity(new_item)
        for ident in idents:
            if "code" in ident: break
        else: entity.add_identifier(new_item,"code",loc)
        existing_xml = entity.metadata_for_entity(new_item,RMNS)
        if existing_xml:
            xml = xml_merge(xml_a=etree.fromstring(existing_xml),xml_b=etree.fromstring(xml))
            entity.update_metadata(new_item,NS,xml)
        else:
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

def xml_merge(xml_a,xml_b):
    for b_child in xml_b.findall('./'):
        a_child = xml_a.find('./' + b_child.tag)
        if a_child is not None:
            if b_child.text:
                a_child.text = b_child.text
        else: print(b_child.tag)
        if b_child.getchildren(): 
            xml_merge(a_child,b_child)
    return etree.tostring(xml_a)

def search_check(search_title,PARENT_REF,desc="",type_flag="Folder",many_flag="first"):
    filters = {"xip.reference":"*","xip.document_type":"*","xip.title":" ""{0}"" ".format(search_title),"xip.parent_ref":f"{PARENT_REF}"}
    search = list(content.search_index_filter_list(query=f"*",filter_values=filters))
    if not len(search):
        SEARCH_REF = PARENT_REF
    elif len(search) > 1:
        print(f'Too many item\'s have been found for: {search_title}; checking if exact match exists by Python...')
        SEARCH_REF = None
        for hit in search:
            if not SEARCH_REF:
                if hit['xip.title'] == search_title:
                    SEARCH_REF = hit['xip.reference']
                    break
            else:
                continue
        if not SEARCH_REF:
            print(f'Couldn\'t find an exact match for: {search_title}, defaulting to fallback: {many_flag}')
            if search[0]['xip.document_type'] == "SO":
                if many_flag == "create":
                    SEARCH_REF = PARENT_REF
                elif many_flag ==  "first":
                    SEARCH_REF = search[0]['xip.reference']
                elif many_flag == "donothing":
                    SEARCH_REF = None
                else: print('Defaulting to "first" option'); SEARCH_REF = search[0].get(['xip.reference'])
                print(f'Reference set as: {SEARCH_REF}')
            elif search[0]['xip.document_type'] == "IO":
                print('Multiple Matches to Items...')
                SEARCH_REF = search[0]['xip.reference'] #Sets SEARCH_REF to None if the Asset Already Exists (Multiple times...).
        else:
            print(f'An exact match found: importing to: {SEARCH_REF}')
            pass
    #Exact hit found. Return result as SEARCH_REF
    else: SEARCH_REF = search[0]['xip.reference']

    #Create Path Scenario.
    #Checks if SEARCH_REF has been assigned as PARENT_REF. 
    if SEARCH_REF == PARENT_REF:
        if type_flag == "Folder":
            SEARCH_REF = preservica_create_folder(search_title,desc,SEARCH_REF)
        elif type_flag == "Asset":
            SEARCH_REF = preservica_create_item(search_title,desc,SEARCH_REF)
    elif SEARCH_REF:
        #Match Scenario (Multi & Exact). No further action required
        pass
    else:
        #None Path Scenario.
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
            box_xml,loc = XML_P.metadata_parse(box_dict=box,item_dict=None,transfer_dict=transfer_dict,type_flag=type_flag,up_flag=up_flag)
            BOX_PRES_REF = search_check(box.get('Box Reference'),PARENT_REF,desc=box.get('Description'))
            if not BOX_PRES_REF: pass # If Search Check returns None (Only on 'donothing' tag), 
            else:
                metadata_update_box(BOX_PRES_REF,box_xml,loc,RMNS)
                if up_flag == "Legal": assigned_retention = None
                else: assigned_retention = box.get('Classification')
                item_list = dfitem.to_dict('records')
                item_search_filters = {"xip.reference":"*","xip.document_type":"IO","xip.title":"*","xip.parent_ref":f"{BOX_PRES_REF}"}
                item_search_dictlist = list(content.search_index_filter_list(query=f"*",filter_values=item_search_filters))
                item_search_list = [x['xip.title'] for x in item_search_dictlist]
                if up_flag == "Legal": field_string = "Agreement Reference"
                elif up_flag == "Lab": field_string = "Title"
                else: field_string = "Item Reference"
                for item in item_list:
                    #for sitem in item_search_list: print(sitem)
                    if any(str(item.get(field_string)) in str(sitem) for sitem in item_search_list):
                        print('Item has a match and is already on Preservica... Skipping') 
                        pass
                    else:
                        if str(box.get('Box Reference')) == str(item.get('Box Verify')):
                            if up_flag == "Legal" or up_flag == "Lab" or up_flag == "USCC": type_flag = up_flag
                            else: type_flag = "Item"
                            item_xml, loc = XML_P.metadata_parse(box_dict=box,item_dict=item,transfer_dict=transfer_dict,type_flag=type_flag,up_flag=up_flag)
                            ITEM_PRES_REF = preservica_create_item(item.get(field_string),item.get('Description'),BOX_PRES_REF)
                            #ITEM_PRES_REF = search_check(itemref,BOX_PRES_REF,desc=itemdesc,type_flag="Asset")
                            #if not ITEM_PRES_REF: pass #Skips Items that are already uploaded onto System.
                            #else: 
                            metadata_update_item(ITEM_PRES_REF,item_xml,loc,RMNS)
                            if str(assigned_retention) != "nan" or assigned_retention != None:
                                policy = retention_lookup(assigned_retention,dfclass)
                                RETENTION_LIST.append({"Preservica Reference":ITEM_PRES_REF,"Preservica Retention":policy})
                            else: print('Retention Not Assigned, skipping policy assignment...')
                        else: 
                            #print('Item does not have a matching Box Verify Reference... Skipping.')
                            pass

def make_structure(folder):
    if os.path.exists(folder): pass
    else: os.makedirs(folder)

if __name__ == '__main__':

    upload_folder = sys.argv[1]
    upload_folder = rf"{upload_folder}"
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
                if transfer_dict.get('Database') == "Certificates": TOP_REF = TMCERT_ROOT
                elif transfer_dict.get('Database') == "Agreements": TOP_REF = TMAGR_ROOT
                else: print('Something Fishy is going on, and I don\'t like it one bit...'); raise SystemExit()
                xl.close()
                box_item_loop(TOP_REF)
            elif "Legal" in xl.sheet_names:
                up_flag = "Legal"
                legal_dept_search = transfer_dict.get('Top-Level')
                if legal_dept_search == "UKI": TOP_REF = UKI_ROOT
                elif legal_dept_search == "UNI": TOP_REF = UNI_ROOT
                else: print('Top-Level needs to be either UNI or UKI; raising System Exit'); raise SystemExit()
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Legal')
                xl.close()
                box_item_loop(TOP_REF)
            #Lab Notebooks will default to 
            elif "Lab" in xl.sheet_names:
                up_flag = "Lab"
                TOP_REF = LAB_ROOT
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Lab')
                dfclass = xl.parse('VAL-CLASS')
                xl.close()
                box_item_loop(TOP_REF)
            else: 
                dfbox = xl.parse('Box')
                dfitem = xl.parse('Item')
                dfclass = xl.parse('VAL-CLASS')
                xl.close()
                up_flag = False
                top_search = transfer_dict.get('Top-Level Folder')
                TOP_REF = search_check(top_search,PRES_ROOT_FOLDER,desc=top_search)
                dept_search = transfer_dict.get('Departmental Folder')
                DEPT_REF = search_check(dept_search,TOP_REF,desc=dept_search)
                date_title = datetime.strftime((transfer_dict.get('Transfer Date')),"%d/%m/%Y")
                date_desc_temp = str(transfer_dict.get('Transfer Description'))
                if date_desc_temp == "nan": date_desc_temp == date_title
                date_desc = date_desc_temp
                DATE_REF = search_check(date_title, DEPT_REF, desc=date_desc)
                box_item_loop(DATE_REF)
            shutil.move(temp_path,comp_path)
            print(f'Successfully moved Temp_File: {temp_path}; to: {comp_path}')
            print(f'Processed Structure Upload, ran for: {datetime.now() - startTime}')
            print(f'Sleeping 10 Secs, before running Retention Assignment...')
            time.sleep(10)
            if RETENTION_LIST:
                retention_assignment(RETENTION_LIST)
            else: 
                print('Retention List is Empty, skipping Retention Assignment...')
        else: 'Ignoring Non-Excel file'
    print(f"Complete!, ran for: {datetime.now() - startTime}")
    if ERROR_LIST:
        dferr = pd.DataFrame(ERROR_LIST,columns=['Pres Ref','Error Message'])
        with pd.ExcelWriter(os.path.join(error_folder,f"Error List_{timestamp}" + ".xlsx"),engine='auto',mode='w') as xlWriter:
            dferr.to_excel(xlWriter,'Errors',index=False)
    if RETENT_ERRORLIST:
        dferr = pd.DataFrame(RETENT_ERRORLIST,columns=['Pres Ref','Error Message'])
        with pd.ExcelWriter(os.path.join(error_folder,f"Retention_Error List_{timestamp}" + ".xlsx"),engine='auto',mode='w') as xlWriter:
            dferr.to_excel(xlWriter,'Errors',index=False)