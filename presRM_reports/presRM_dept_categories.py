from pyPreservica import *
import pandas as pd
from lxml import etree
from xml.sax.saxutils import escape
import secret

entity = EntityAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")

filepath = r"C:\Users\Chris.Prince\Downloads\Cat_Dept_Output.xlsx"

df = pd.read_excel(filepath,index_col=0)

new_list = df[['Preservica Reference','dptnotes','dptfrmname','dpthistnotes']].values.tolist()

RMNS = "http://rm.unilever.co.uk/schema"

def xml_add(xml_a,xml_b,x_child=None):
    for b_child in xml_b.findall('./'):    
        a_child = xml_a.find('./' + b_child.tag)
        if a_child is None:
            a_child = etree.SubElement(x_child,b_child.tag)
            a_child.text = b_child.text
        else:
            x_child = a_child
            xml_add(a_child,b_child,x_child)
    return etree.tostring(xml_a)

def text_format(x):
    if str(x) == "nan": return escape(str(x).replace("nan",""))
    else: return escape(str(x))

for x in new_list:
    print(f"Processing: {x[0]}")
    dept_folder = entity.folder(x[0])
    xml_pres_s = entity.metadata_for_entity(dept_folder,RMNS)
    xml_pres = etree.fromstring(xml_pres_s)
    
    x[1] = text_format(x[1])
    x[2] = text_format(x[2])
    x[3] = text_format(x[3])

    xml_new_s =f"""<rm:rm xmlns:rm="{RMNS}">
    <rm:legacyinfo>
    <rm:departmentnotes>{str(x[1])}</rm:departmentnotes>
    <rm:departmentformername>{str(x[2])}</rm:departmentformername>
    <rm:departmenthistoricnotes>{str(x[3])}</rm:departmenthistoricnotes>
    </rm:legacyinfo>
    </rm:rm>"""

    xml_new = etree.fromstring(xml_new_s)
    xml_upload = xml_add(xml_pres,xml_new)
    entity.update_metadata(dept_folder,RMNS,data=xml_upload.decode('utf-8'))