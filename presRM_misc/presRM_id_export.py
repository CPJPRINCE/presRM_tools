import pandas as pd
from pyPreservica import *
from .. import secret
import datetime
from lxml import etree

startTime = datetime.datetime.now()
print("Start Time: " + str(startTime))

entity = EntityAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

retention = RetentionAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

if __name__ == '__main__':
    uknl = "UK"
    
    if uknl == "UK":
        nl_flag = False
    elif uknl == "NL":
        nl_flag = True
    else:
        uknl == "UK"
    if nl_flag:  rootname = "Records Management NL"
    else: rootname = "Records Management UK" # Can be overriden with sub-childs for testing...

    xlpath = rf"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Records Management\Project Snakeboot\XLoutput\CS10Export_{uknl}_Final.xlsx"

    #Lists for Error Checking and for Retention Setting.
    RETENT_ERRORLIST = []
    RETENTION_LIST = []
    #Root Preservica Level, where database will sit. Hard coded, doesn't need to vary - only update on final upload
    PRES_ROOT_FOLDER = "bf614a5d-10db-4ff8-addb-2c913b3149eb"

    #Data Frame Loading. Sheets needs to be changed on changing from UK to NL. Normally takes ~2 minutes to read UK.
    
    filters = {"xip.parent_hierarchy":PRES_ROOT_FOLDER,'rm.legacyid':"","xip.reference": ""}
    content.search_callback(content.ReportProgressCallBack())
    content.search_index_filter_csv("%", "pres_RM-UK.csv", filter_values=filters)
