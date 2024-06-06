from pyPreservica import ContentAPI
import time
import os
try:
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    try: 
        content = ContentAPI(username=input("Please enter your Username: "),password=input("Please enter your Password: "),server="unilever.preservica.com")
    except:
        print("Wrong Credentials please try again")
        content = ContentAPI(username=input("Please enter your Username: "),password=input("Please enter your Password: "),server="unilever.preservica.com")

    filters = {}
    TMAGREEDB = "b4a34349-f5d8-4096-bd44-8a2a30e7482c"
    TMCERTDB = "07c847df-f708-408b-b7d4-0ed4ae48de31"
    content.search_callback(content.ReportProgressCallBack())

    def define_filters(REF):
        filters = {"xip.parent_ref":REF,"xip.title":"","xip.description":"","tm.location":"","tm.box":"","tm.country":"","tm.trademark":"","tm.registration":"","tm.type":"","tm.parties":"","tm.add":""}
        return filters
    
    filters = define_filters(TMAGREEDB)
    search = content.search_index_filter_csv("%","TMAGREEDB.csv",filter_values=filters)
    filters = define_filters(TMCERTDB)
    search = content.search_index_filter_csv("%","TMCERTDB.csv",filter_values=filters)
except Exception as e:
    print(e)
    time.sleep(10)