import pandas as pd
from pyPreservica import *
import secret
import re
from datetime import datetime
from lxml import etree
import html

startTime = datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    tenant="UARM",server="unilever.preservica.com")

def lookup_loop(folder_ref):
    children = list(entity.descendants(folder_ref))
    if len(children) == 0:
        print("Empty Dir!")
        pass
    else:
        for e in children:
            print(e.title)
            print(e.entity_type)
            if e.entity_type == "EntityType.ASSET": 
                print(e.entity_type, e.reference)
                lookup_loop(folder_ref)
                pass
            else: lookup_loop(e.reference)

if __name__ == '__main__':

    PRES_ROOT_FOLDER = "09114261-3242-4afa-b569-4f32008642bc"
    PRES_ROOT_FOLDER = None
    lookup_loop(PRES_ROOT_FOLDER)
    print('Complete!')
    print(f"Ran for: {datetime.now() - startTime}")
    