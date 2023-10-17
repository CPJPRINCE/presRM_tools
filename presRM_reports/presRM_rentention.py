from pyPreservica import *
from .. import secret
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

print(pyPreservica.__version__)

startTime = datetime.now()
print(f"Start time: {str(startTime)}")

content = ContentAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")

entity = EntityAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")

#print(content.indexed_fields())
#print(entity.token)

def report_expired(assignments):
    for a in assignments:
        if a.expired:
            retentionList.append([ref,f['xip.title'],pol_name,a.expired,a.start_date])
    return retentionList
    
def report_not_expired(assignments):
    for a in assignments:
        due_date = datetime.year(a.period) + datetime(a.start_date)
        if a.expired:
            retentionList.append([ref,f['xip.title'],pol_name,a.expired,a.start_date,a.period,due_date])
    return retentionList
    pass

if __name__ == '__main__':

    TOP_REF = "cd830d4f-a41f-4610-ba89-9bcd3de52e51"

    output_directory = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Archive\Digital Preservation\github\cs10_reports"

    #REPORT = "ARCHIVE"
    REPORT = "CHECK"
    #REPORT = "DESTROY"
    EXPIRED_CHECK = True

    metadata_fields = {
        "xip.reference": "*",
        "xip.title": "*",
        "xip.description": "*",
        "xip.security_descriptor": "*",
        "xip.identifier": "*",
        "xip.parent_hierarchy": TOP_REF,
        "xip.retention_policy_assignment_ref": "*",
        "xip.document_type": "io",
        "xip.retention_policy_assignment_name": REPORT,
        "rm.location": "*",
        "rm.defaultlocation": "*",
        "rm.client": "*"
        }
    retentionList = []
    report_search = list(content.search_index_filter_list(query="%", filter_values=metadata_fields))
    t = len(report_search)
    #print(report_search)
    print(f"Total Results: {str(t)}")
    if not t == 0:
        c = 0
        for f in report_search:
            #print(f)
            c += 1
            print(f"Processing: {c} / {t}", end="\r")
            ref = f['xip.reference']
            #print(ref)
            title = f['xip.title']
            desc = f['xip.description']
            parent = f['xip.parent_ref']
            pol_name = f['xip.retention_policy_assignment_name'][0]
            pol_ref = f['xip.retention_policy_assignment_ref'][0]
            boxloc = f['rm.location']
            boxdloc = f['rm.defaultlocation']
            client = f['rm.client']
            #print(boxloc)
            #print(entity(ref))
            asset = entity.asset(ref)
            #print(pol_ref)
            poly = retention.policy(reference=pol_ref)
            asset = entity.asset(ref)
            #print(asset)
            assg = retention.assignments(asset)
            for a in assg:
                expir = a.expired
                startd = a.start_date
            #print(poly.period)
            #print(startd)
            due_date = datetime.strptime(startd,"%Y-%m-%dT%H:%M:%SZ") + relativedelta(days=int(poly.period))
            xiplist = f['xip.parent_hierarchy'][0].split()
            x = entity.asset(ref)
            fpath = x.title
            f = entity.folder(x.parent)
            parent_name = f.title
            while f.parent is not None:
                f = entity.folder(f.parent)
                fpath = f.title + ":::" + fpath
            fpath = fpath.rsplit(":::",maxsplit=1)[0]
            fpath = fpath.split(":::",maxsplit=2)[-1]
            url = f"https://unilever.preservica.com/explorer/explorer.html#browse:IO&{ref}"
            if EXPIRED_CHECK: 
                if expir:
                    retentionList.append([ref,url,fpath,parent,parent_name,
                                        title,desc,boxloc,
                                        boxdloc,client,pol_name,
                                        pol_ref,startd,poly.period,
                                        due_date,expir])

                else:
                    print("Retention Not Assigned? Why is appearing here...")
                    pass
            else:
                retentionList.append([ref,url,fpath,
                                        title,desc,boxloc,
                                        boxdloc,client,pol_name,
                                        pol_ref,startd,poly.period,
                                        due_date,expir])
                
        df = pd.DataFrame(retentionList,columns=["Preservica Reference","URL","Parent Hierarchy","Preservica Parent Reference","Parent Title"
                                                 "Title", "Description", "Location",
                                                 "Home Location","Client", "Policy Name",
                                                 "Policy Reference","Retention Start Date", "Retention Period",
                                                 "Finish Date","Expired Flag"])
        
        # boxlist = df['Parent Title'].drop_duplicates().values.tolist()
        # itemlist = []
        # for box in boxlist:
        #     idx = df.index[df['Parent Title'] == box].tolist()
        #     for i in idx:
        #         row = df.loc[i]
        #         itemlist.append(row['Title'])
        #         set(itemlist)

                
        output_file = os.path.join(output_directory,f'RetentionReport-{REPORT + "_" + str(datetime.now().strftime("%Y-%m-%d"))}.xlsx')
        with pd.ExcelWriter(output_file,engine='auto',mode='w') as writer:
            df.to_excel(writer,index=False,sheet_name="Items")
            #boxdf.to_excel(writer,index=False,sheet_name="Box")
        print(f'File Saved to: {output_file}')
        print(f"Complete, running time: {datetime.now() - startTime}")
    else: print('No Results Return')
    