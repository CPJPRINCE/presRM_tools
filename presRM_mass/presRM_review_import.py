from pyPreservica import *
import pandas as pd
import sys
from lxml import etree as ET
import secret

file = sys.argv[1]

df = pd.read_excel(file, 'finalreviews')
df = df.fillna("")
records_dict = df.to_dict('records')

entity = EntityAPI(username=secret.username, password=secret.password, server=secret.server)

reviewns = "http://rm.unilever.co.uk/review/v1"

for record in records_dict[:2]:
    reference = record.get('reference')
    type = record.get('document_type')
    xml_root = ET.Element(f"{{{reviewns}}}rmreview", nsmap={"rmreview":reviewns})
    for k,v in record.items():
        if k == "reference" or k == "document_type" or k == "title":
            pass
        else:
            ET.SubElement(xml_root, f"{{{reviewns}}}{k}").text = str(v)
    if type == "IO":
        e = entity.asset(reference)
    elif type == "SO":
        e = entity.folder(reference)
    entity.add_metadata(e, reviewns, ET.tostring(xml_root).decode("utf-8"))
    print(f"Processing: {reference}")