import csv
from pyPreservica import *
import secret
from lxml import etree

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")
csvpath = r"C:\Users\Chris.Prince\Downloads\Artificial Files.csv"

RMNS = "http://rm.unilever.co.uk/schema"


with open(csvpath,'r') as w:
    c = csv.DictReader(w)
    for number,row in enumerate(c):
        ref = row.get('Entity Ref')
        print(ref)
        asset = entity.asset(ref)
        m = entity.metadata_for_entity(asset,RMNS)
        xml_str = etree.fromstring(m)
        obj = xml_str.find(f'.//{{{RMNS}}}objtype')
        obj.text = "Item"
        xml_nstr = etree.tostring(xml_str).decode('utf-8')
        entity.update_metadata(asset,RMNS,xml_nstr)