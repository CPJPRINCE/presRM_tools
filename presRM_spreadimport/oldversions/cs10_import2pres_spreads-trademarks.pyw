import pandas as pd
from pyPreservica import *
import secret
import re
import shutil

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

def metadata_update_item(PRES_REF,locator, xml):
    asset = client.asset(PRES_REF)
    client.add_identifier(asset,"code",locator)
    client.add_metadata(asset,TMNS,xml)
    
def define_metadata(metadata):
    title = str(metadata[1]) + "/" + str(metadata[2] + " - " + str(metadata[4]))
    description = str(metadata[3] + " - " + str(metadata[5]) + " - " + str(metadata[6]))
    code = str(metadata[0])
    country = str(metadata[3])
    box =  str(metadata[1])
    return [title,description,code,country,box]

def remove_nan(record):
    if record == "nan":
        record = "nan"
def xml_generate(record):
    ##Customise to remove all special characters...
    record = [str(rec).replace("&","&amp;") for rec in record]
    Location = remove_nan(record[0])
    Box = remove_nan(record[1])
    Reference = remove_nan(record[2])
    Country = remove_nan(record[3])
    Trademark = remove_nan(record[4])
    Registration = remove_nan(record[5])
    Type = remove_nan(record[6])
    Parties = remove_nan(record[7])
    Add = remove_nan(record[8])
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
    return xml,[Location,Box,Reference,Country,Trademark,Registration,Type,Parties]

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

def create_above_branch(choice,xml,metadata):
    if choice == "Box":
        branch = metadata[4]
    elif choice == "Country":
        branch = metadata[3]
    else: branch = None, print('Branch needs to be set to valid option!')
    if branch:
        PRES_DEST = search_for_folder(branch,TOPLEVEL_FOLDER_REFERENCE)
        if PRES_DEST == "CREATEFOLDER":
            SUB_PRES_DEST = preservica_create_folder(branch,branch,TOPLEVEL_FOLDER_REFERENCE)
            print(f'New Folder Created: {branch}, waiting 10s to reset search index')
            for x in range(10,1,-1):
                time.sleep(1)
                print(x,end='\r',flush=True)
            new_item_ref = preservica_create_item(metadata[0],metadata[1],SUB_PRES_DEST)
            metadata_update_item(new_item_ref,metadata[2],xml)
            print(f'Created Item at: {new_item_ref}')
        elif PRES_DEST == "ERROR - TOO MANY HITS":
            print('Folder will not be created skipping import for...')
            pass
        else:
            new_item_ref = preservica_create_item(metadata[0],metadata[1],PRES_DEST)
            metadata_update_item(new_item_ref,metadata[2],xml)
            print(f'Created Item at: {new_item_ref}')
            pass
    else: print('Missing Metadata from branch... It needs to be set!')

def xl_import_2_preservica(method,df):
    dflist = df.values.tolist()
    for record in dflist:
        xml,metalist = xml_generate(record)
        metadata = define_metadata(metalist)
        if method == "Flat":
            new_item_ref = preservica_create_item(metadata[0],metadata[1],TOPLEVEL_FOLDER_REFERENCE)
            metadata_update_item(new_item_ref,metadata[2],xml)
            print(f'Created Item at: {new_item_ref}')
            pass
        elif method == "Country":
            create_above_branch("Country",xml,metadata)
            pass
        elif method == "Box":
            create_above_branch("Box",xml,metadata)
            pass

def test_read(presref):
    folder = presref
    for url in folder.metadata:
        xml_data = client.metadata(url) 
        print(xml_data)

def test_upload(presref):
    xmltest = f"<tm:tm xmlns:tm='http://rm.unilever.co.uk/trademarks/v1' xmlns='http://rm.unilever.co.uk/trademarks/v1'>" \
        f"<tm:location>Location</tm:location>" \
        f"<tm:box>Box</tm:box>" \
        f"<tm:reference>Reference</tm:reference>" \
        f"<tm:country>Country</tm:country>" \
        f"<tm:trademark>poopybuttpants</tm:trademark>" \
        f"<tm:registration>Registration</tm:registration>" \
        f"<tm:type>Type</tm:type>" \
        f"<tm:parties>Parties</tm:parties>" \
        f"<tm:add>?</tm:add>" \
        "</tm:tm>"
    folder = client.asset(presref)
    folder = client.add_metadata(folder,"http://rm.unilever.co.uk/trademarks/v1",xmltest)
    print('Complete')

if __name__ == '__main__':

    method = "Flat"
    TMNS = "http://rm.unilever.co.uk/trademarks/v1"
    upload_folder = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Archive\Digital Preservation\github\cs10_extraction\Uploads\Trademarks"
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
            shutil.move(file_path,temp_path)
            dfdataselection = pd.read_excel(temp_path,"TOP")
            dftm = pd.read_excel(temp_path,"TM")
            destination = dfdataselection['Database'].values.item()
            if destination == "Certificates":
                TOPLEVEL_FOLDER_REFERENCE = "b37bf77e-05d9-4adc-a41f-534a94330582"
            elif destination == "Agreements":
                TOPLEVEL_FOLDER_REFERENCE = "621be7cf-245a-4a03-93ae-b368251e6f76"
            else:
                print('Error, TOP must either be Agreements or Certificates')
                break
            xl_import_2_preservica(method,dftm)
            shutil.move(temp_path,comp_path)
        else: print('Ignoring non-Excel file')