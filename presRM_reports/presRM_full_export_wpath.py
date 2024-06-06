import pandas as pd
from pyPreservica import *
import secret
import datetime
from lxml import etree
import pandas as pd

startTime = datetime.datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

def path_parent_return(ref,io):
    if io == "IO": x = entity.asset(ref)
    elif io == "SO": x = entity.folder(ref)
    else: print("Error?")
    fpath = x.title        
    f = entity.folder(x.parent)
    parent_name = f.title
    while f.parent is not None:
        f = entity.folder(f.parent)
        fpath = f.title + ":::" + fpath
    if PATH_FLAG:
        fpath = fpath.rsplit(":::",maxsplit=1)[0]
        fpath = fpath.split(":::",maxsplit=2)[-1]
    return fpath,parent_name

def checksum_return(ref,io):
    if io =="IO":
        x = entity.asset(ref)
        fixalg = []
        fixity = []
        try: 
            for r in entity.representations(x):
                for c in entity.content_objects(r):
                    for g in entity.generations(c):
                        for b in g.bitstreams:
                                for alg,hash in b.fixity.items():
                                    fixalg.append(alg)
                                    fixity.append(hash) 
        except Exception as e:
            fixalg = "Error"
            fixity = "Error"
    else: 
        fixalg = ""
        fixity = ""
    return fixalg,fixity

if __name__ == '__main__':

    PRES_ROOT_FOLDER = sys.argv[1]
    SPLIT_FLAG = False
    PATH_FLAG = False
    HASH_FLAG = False
    filters = {"xip.parent_hierarchy":PRES_ROOT_FOLDER,
               "xip.title": "",
               "xip.description":"",
               "xip.document_type":"",
               "rm.statusdate":"",
               "rm.coverdate":"",
               "rm.notes":"",
               "rm.disposition":"",
               "rm.objtype":"",
               "rm.boxtype":"",
               "rm.location":"",
               "rm.weight":"",
               "rm.area":"",
               "rm.defaultlocation":"",
               "rm.client":"",
               "rm.alternatecontact":"",
               "rm.itemtype":"",
               "rm.format":"",
               "rm.numberofitems":"",
               "rm.transferdate":"",
               "rm.transferacceptedby":"",
               "rm.legacyid":"",
               "rm.legacyparentid":""}
    dict_list = []
    content.search_callback(content.ReportProgressCallBack())
    
    report_search = list(content.search_index_filter_list(query="%", filter_values=filters))
    tot = len(report_search)
    c = 0
    for hit in report_search:
        c += 1
        print(f'Processing {c} / {tot}',end="\r")
        if PATH_FLAG:
            full_path,parent_title = path_parent_return(hit.get('xip.reference'),hit.get('xip.document_type'))
            hit.update({'Path': full_path, 'Parent':parent_title})
        if HASH_FLAG:
            alg, hash = checksum_return(hit.get('xip.reference'),hit.get('xip.document_type'))
            hit.update({'Algorithm': alg,'Fixity': hash})
        dict_list.append(hit)
    df = pd.DataFrame.from_records(dict_list)
    df.index.name = "Index"
    df.to_excel('Output.xlsx')