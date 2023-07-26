import pandas as pd
from pyPreservica import *
import secret
from datetime import datetime 

#print(pyPreservica.__file__)

reportlocation = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Archive\Digital Preservation\github\cs10_reports"

content = ContentAPI(username=secret.username,password=secret.password, \
                    tenant="UARM",server="unilever.preservica.com")

alist = []

starttime = datetime.now()
print(f"Started: {starttime}")
top_ref = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
objtype = "Box"

filters = {"xip.reference":"*","xip.identifier":"code","xip.parent_hierarchy":top_ref,"rm.objtype":"Box","rm.boxtype":"*",}
c = 0
content.search_callback(content.ReportProgressCallBack())
location_search = list(content.search_index_filter_list(query="%",filter_values=filters))
tot = len(location_search)
for f in location_search:
    x = f['xip.identifier'][0].split(" ")
    alist.append([f['xip.reference'], x[1]])
    c +=1
    print(f"Processing A List: {c} / {tot}",end="\r")
print()       

dfa = pd.DataFrame(alist,columns = ['Reference','Location'])

nlist = []
for n in range(1,109):
    for m in range(1,181):
        nlist.append(str(n) + ":" + str(m).zfill(3))

c = 0
tot = len(nlist)
totlist = nlist
print('Processing N List...')
nlist = filter(lambda n: dfa[dfa['Location'] == n].empty,nlist)

print()       

dfn = pd.DataFrame(nlist,columns = ['Free Locations'])
dftot = pd.DataFrame(totlist,columns = ['All Locations'])
#output_file = os.path.join(reportlocation,'Locations-Report_' + datetime.today().strftime("%Y-%m-%d") + ".xlsx")
output_file = os.path.join(reportlocation,'LocationsDB.xlsx')
with pd.ExcelWriter(output_file) as w:   
    dfn.to_excel(w,sheet_name='Free',index=False)
    dfa.to_excel(w,sheet_name='Used',index=False)
    dftot.to_excel(w,sheet_name='All',index=False)
print(f"Finished, ran for: {datetime.now() - starttime}")