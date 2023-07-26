import pandas as pd
from pyPreservica import *
import secret
import re
import shutil
from lxml import etree

client = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")


def preservica_create_folder(title,description,PRES_PARENT_REF):
    new_folder = client.create_folder(title,description,"closed",PRES_PARENT_REF)
    return new_folder.reference

def preservica_create_item(title,description,PRES_PARENT_REF):
    parent = client.folder(PRES_PARENT_REF)
    new_item = client.add_physical_asset(title,description,parent,"closed")
    return new_item.reference

def metadata_update_item(PRES_REF,locator,xml):
    asset = client.asset(PRES_REF)
    client.add_identifier(asset,"code",locator)
    client.add_metadata(asset,TMNS,xml)
    

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

def metadata_parse(record):
    ##Customise to remove all special characters...
    Location = parse_text(record.get('Location'))
    Box = parse_text(record.get('Box'))
    Reference = parse_text(record.get('Box'))
    Country = parse_text(record.get('Country'))
    Trademark = parse_text(record.get('Trademark Name'))
    Registration = parse_text(record.get('Registration Number'))
    Type = parse_text(record.get('Type of Document'))
    Parties = parse_text(record.get('Parties Involved'))
    Add = parse_text(record.get('Additional Information'))
    xml = f"<tm:tm xmlns:tm='{TMNS}' xmlns='{TMNS}'>" \
        f"<tm:location>{Location}</tm:location>" \
        f"<tm:box>{Box}</tm:box>" \
        f"<tm:reference>{Reference}</tm:reference>" \
        f"<tm:country>{Country}</tm:country>" \
        f"<tm:trademark>{Trademark}</tm:trademark>" \
        f"<tm:registration>{Registration}</tm:registration>" \
        f"<tm:type>{Type}</tm:type>" \
        f"<tm:parties>{Parties}</tm:parties>" \
        f"<tm:add>{Add}</tm:add>" \
        "</tm:tm>"

    try: xml_test = etree.fromstring(xml)
    except Exception as e: 
        print(e)
        print(xml)
    
    return xml
### Below two Functions are only used in methodology that are not 'Flat'
# Flat is only method being used.
# Left these in script for reference purposes. 

def search_for_folder(folder_title,parent):
    filters = {"xip.reference":"*","xip.title":f"'{folder_title}'","xip.parent_ref":f"{parent}"}
    search = list(content.search_index_filter_list(query=f"{folder_title}",filter_values=filters))
    if not len(search):
        print('No Hits')
        dest_folder = "CREATEFOLDER"
    elif len(search) > 1:
        print('Too many')
        dest_folder = "ERROR - TOO MANY HITS"
    else:
        print('Just right')
        dest_folder = search[0]['xip.reference']
    return dest_folder

def parse_title_and_desc(record):
    Title = str(record.get('Box')) + " - " + str(record.get('Reference')) + " - " + str(record.get('Trademark Name'))
    Description = str(record.get('Country')) + " - " + str(record.get('Registration Number')) + " - " + str(record.get('Type of Document'))
    return Title, Description

#Unused Function
def create_above_branch(choice,xml,record):
    Title, Description = parse_title_and_desc(record)
    if choice == "Box":
        branch = record.get('Box')
    elif choice == "Country":
        branch = record.get('Country')
    else: branch = None, print('Branch needs to be set to valid option!')
    if branch:
        PRES_DEST = search_for_folder(branch,TOPLEVEL_FOLDER_REFERENCE)
        if PRES_DEST == "CREATEFOLDER":
            SUB_PRES_DEST = preservica_create_folder(branch,branch,TOPLEVEL_FOLDER_REFERENCE)
            print(f'New Folder Created: {branch}, waiting 10s to reset search index')
            for x in range(10,1,-1):
                time.sleep(1)
                print(x,end='\r',flush=True)
            new_item_ref = preservica_create_item(Title,Description,SUB_PRES_DEST)
            metadata_update_item(new_item_ref,record.get('Location'),xml)
            print(f'Created Item at: {new_item_ref}')
        elif PRES_DEST == "ERROR - TOO MANY HITS":
            print('Folder will not be created skipping import for...')
            pass
        else:
            new_item_ref = preservica_create_item(Title,Description,PRES_DEST)
            metadata_update_item(new_item_ref,record.get('Location'),xml)
            print(f'Created Item at: {new_item_ref}')
            pass
    else: print('Missing Metadata from branch... It needs to be set!')

def xl_import_2_preservica(method,df):
    record_dict = df.to_dict('records')
    for record in record_dict:
        xml = metadata_parse(record)

        Title,Description = parse_title_and_desc(record)
        if method == "Flat":
            new_item_ref = preservica_create_item(Title,Description,TOPLEVEL_FOLDER_REFERENCE)
            metadata_update_item(new_item_ref,record.get('Location'),xml)
            print(f'Created Item at: {new_item_ref}')
            pass
        elif method == "Country":
            create_above_branch("Country",xml,record)
            pass
        elif method == "Box":
            create_above_branch("Box",xml,record)
            pass

if __name__ == '__main__':

    method = "Flat"
    TMNS = "http://rm.unilever.co.uk/trademarks/v1"
    upload_folder = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Preservica Spreadsheet Uploads\Trademarks - PLACE COMPLETED TEMPLATE IN HERE"
    temp_folder = os.path.join(upload_folder,"temp")
    if os.path.exists(temp_folder): pass
    else: os.makedirs(temp_folder)
    comp_folder = os.path.join(upload_folder,"complete")
    if os.path.exists(comp_folder): pass
    else: os.makedirs(comp_folder)

    for file in os.listdir(upload_folder):
        file_path = os.path.join(upload_folder,file)
        temp_path = os.path.join(temp_folder,file)
        comp_path = os.path.join(comp_folder,file)
        if os.path.isfile(file_path) and file_path.endswith('.xlsx'):
            print(f'Temp Move Complete: {temp_path}')
            shutil.move(file_path,temp_path)
            dfdataselection = pd.read_excel(temp_path,"TOP")
            dftm = pd.read_excel(temp_path,"TM")
            destination = dfdataselection['Database'].values.item()
            if destination == "Certificates":
                TOPLEVEL_FOLDER_REFERENCE = "07c847df-f708-408b-b7d4-0ed4ae48de31"
            elif destination == "Agreements":
                TOPLEVEL_FOLDER_REFERENCE = "b4a34349-f5d8-4096-bd44-8a2a30e7482c"
            else:
                print('Error, TOP must either be Agreements or Certificates')
                break
            xl_import_2_preservica(method,dftm)
            shutil.move(temp_path,comp_path)
            print(f'Move Complete: {comp_path}')
        else: print('Ignoring non-Excel file')