from pyPreservica import *
import pandas as pd
import secret
from lxml import etree

content = ContentAPI(username=secret.username, password=secret.password,server="unilever.preservica.com")
entity = EntityAPI(username=secret.username, password=secret.password,server="unilever.preservica.com")
#file_path = sys.argv[1]

file_path = r"C:\Users\Chris.Prince\Downloads\Chris - Destructions Shelf Locs for Removal from Preservica.xlsx"

df = pd.read_excel(file_path)

print(df)

name_list = df['Name'].values.tolist()

pheir = "3cea8f86-1f4e-43dc-81c3-cfa81d5370f0"

RMNS = "http://rm.unilever.co.uk/schema"

def main(folder_name):
    filters = {"xip.parent_hierarchy": pheir,"xip.title": f" ""{}"" ".format(folder_name),"xip.document_type":"SO"}
    search = list(content.search_index_filter_list("%",filter_values=filters))
    if len(search) == 0: print(f'No Hits for {folder_name}')
    elif len(search) > 1: print(f'Multiple Matches for {folder_name}')
    else:
        for hit in search[0:2]:
            print(hit.get('xip.title'),hit.get('xip.reference'))
            entity_update_folder(reference=hit.get('xip.reference'))
            for asset in entity.descendants(hit.get('xip.reference')):
                print(asset.title, asset.reference)
                entity_update_asset(asset.reference)

def entity_update_folder(reference):
    folder = entity.folder(reference)
    emeta = entity.metadata_for_entity(folder,RMNS)
    identifiers = entity.identifiers_for_entity(folder)
    location_check = False
    for i in identifiers:
        if "code" in i:
            location_check = True
    if location_check: entity.delete_identifiers(folder,identifier_type="code")
    xml_data = etree.fromstring(emeta)
    if xml_data.find(f'./{{{RMNS}}}boxinfo/{{{RMNS}}}boxlocation') is not None:
        box_loc = xml_data.find(f'./{{{RMNS}}}boxinfo/{{{RMNS}}}boxlocation')
        box_loc.text = ""
        xml_string = etree.tostring(xml_data).decode('utf-8')
        entity.update_metadata(folder, RMNS, xml_string)
    else: box_loc = None
    return folder

def entity_update_asset(reference):
    asset = entity.asset(reference)
    emeta = entity.metadata_for_entity(asset,RMNS)
    identifiers = entity.identifiers_for_entity(asset)
    location_check = False
    for i in identifiers:
        if "code" in i:
            location_check = True
    if location_check:
        entity.delete_identifiers(asset,identifier_type="code")
    xml_data = etree.fromstring(emeta)
    if xml_data.find(f'./{{{RMNS}}}boxinfo/{{{RMNS}}}boxlocation') is not None:
        box_loc = xml_data.find(f'./{{{RMNS}}}boxinfo/{{{RMNS}}}boxlocation')
        box_loc.text = ""
        xml_string = etree.tostring(xml_data).decode('utf-8')
        entity.update_metadata(asset, RMNS, xml_string)
    else: box_loc = None

for name in name_list:
    main(name)
