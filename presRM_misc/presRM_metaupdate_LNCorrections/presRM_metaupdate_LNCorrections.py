import pandas as pd
from pyPreservica import *
import secret
import datetime
from lxml import etree
from  xml_parsing import XML_Parse


startTime = datetime.datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

def date_formatting(vardate):
    vardate = str(vardate)
    if not vardate or vardate == "False" or vardate == "?" or vardate == "NaN" or vardate == "nan" or vardate == "NaT" or vardate == "": vardate = ""
    else:
        try:
            vardate = datetime.datetime.strptime(vardate,"%Y-%m-%d %H:%M:%S")
            vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except Exception: 
            try: 
                vardate = datetime.datetime.strptime(vardate,"%Y-%m-%dT%H:%M:%S")
                vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
            except Exception:
                try: 
                    vardate = datetime.datetime.strptime(vardate,"%d-%m-%Y")
                    vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                except Exception:
                    vardate = datetime.datetime.strptime(vardate,"%d/%m/%Y")
                    vardate = vardate.strftime("%Y-%m-%dT%H:%M:%S.000Z")
                    vardate = ""
    return vardate

def metadata_update_folder(PRES_REF,xml_string,loc,NS):
    try:
        new_box = entity.folder(PRES_REF)
        if new_box.title != title or new_box.description != description:
            if new_box.title != title:
                new_box.title = title
            elif new_box.title != description:
                new_box.description = description
            entity.save(new_box)
        idents = entity.identifiers_for_entity(new_box)
        for ident in idents:
            if "code" in ident: break    
        else: 
            entity.add_identifier(new_box,"code",loc)
        existing_xml = entity.metadata_for_entity(new_box,NS)
        if existing_xml:
            pass #pass due to error on upload
            #xml_string = xml_merge(xml_a=etree.fromstring(existing_xml),xml_b=etree.fromstring(xml_string))
            #entity.update_metadata(new_box,NS,xml_string.decode('utf-8'))
        else: 
            entity.add_metadata(new_box,NS,xml_string.decode('utf-8'))
    except Exception as e:
        ERROR_LIST.append([PRES_REF,e])
        print(e)

def metadata_update_item(PRES_REF,xml_string,loc,NS):
    try:
        new_item = entity.asset(PRES_REF)
        if new_item.title != title or new_item.description != description:
            if new_item.title != title:
                new_item.title = title
            elif new_item.title != description:
                new_item.description = description
            entity.save(new_item)
        idents = entity.identifiers_for_entity(new_item)
        for ident in idents:
            if "code" in ident: break
        else:
            entity.add_identifier(new_item,"code",loc)
        existing_xml = entity.metadata_for_entity(new_item,NS)
        if existing_xml:
            xml_string = xml_merge(xml_a=etree.fromstring(existing_xml),xml_b=etree.fromstring(xml_string))
            entity.update_metadata(new_item,NS,xml_string.decode('utf-8'))
        else:
            entity.add_metadata(new_item,NS,xml_string.decode('utf-8'))
    except Exception as e:
        ERROR_LIST.append([PRES_REF,e])
        print(e)

def xml_merge(xml_a,xml_b):
    for b_child in xml_b.findall('./'):
        a_child = xml_a.find('./' + b_child.tag)
        if a_child is not None:
            if b_child.text:
                a_child.text = b_child.text
        else:
            #print(b_child.tag)
            pass
        if b_child.getchildren(): 
            xml_merge(a_child,b_child)
    return etree.tostring(xml_a)

if __name__ == '__main__':
    
    ERROR_LIST = []

    RMNS = "http://rm.unilever.co.uk/schema"
    xlpath = rf"C:\Users\Chris.Prince\OneDrive - Unilever\LN_Output_Cleaned_Correct.xlsx"

    if xlpath.endswith('.xlsx'):
        df = pd.read_excel(xlpath,sheet_name='RMUpdate')

    row_list = df.to_dict('records')
    XML_P = XML_Parse(RMNS)
    for row in row_list:
        pres_ref = row.get('Preservica Reference')
        type_flag = row.get('Type Flag')
        title = row.get('Title')
        description = row.get('Description')
        xml_string,loc = XML_P.metadata_parse(box_dict=row,item_dict=row,transfer_dict=row,type_flag=type_flag)
        if type_flag == "Box": 
            metadata_update_folder(pres_ref,xml_string,loc,RMNS)
        elif type_flag == "Item":
            metadata_update_item(pres_ref,xml_string,loc,RMNS)
    print(ERROR_LIST)