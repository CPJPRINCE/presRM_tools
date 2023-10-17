from pyPreservica import *
import secret
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


entity = EntityAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")

retention = RetentionAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")        

content = ContentAPI(username={secret.username},password={secret.password}, \
                    server="unilever.preservica.com")

def parent_return(ref,io):
    if io == "IO": x = entity.asset(ref)
    elif io == "SO": x = entity.folder(ref)
    else: print("Error?")
    while x.parent is not None:
        f = entity.folder(x.parent)
        parent_name = f.title
    return parent_name

def path_return(ref,io,split_flag):
    if io == "IO": x = entity.asset(ref)
    elif io == "SO": x = entity.folder(ref)
    else: print("Error?")
    fpath = x.title
    f = entity.folder(x.parent)
    parent_name = f.title    
    while f.parent is not None:        
        fpath = f.title + ":::" + fpath
        f = entity.folder(f.parent)    
    if split_flag:
        fpath = fpath.rsplit(":::",maxsplit=1)[0]
        fpath = fpath.split(":::",maxsplit=3)[-1]
    return fpath,parent_name

class RetentionReport():
    def __init__(self,REF="bf614a5d-10db-4ff8-addb-2c913b3149eb",OUTPUT_DIR=os.getcwd(),REPORT_TYPE="ALL",EXPIRED_FLAG=True,PATH_FLAG=False,SPLIT_FLAG=True):
         self.reference = REF 
         self.output_dir = os.path.abspath(OUTPUT_DIR)
         self.report_type = REPORT_TYPE.upper()
         self.report_mode = self.report_mode_select()
         self.expired_flag = EXPIRED_FLAG
         self.path_flag = PATH_FLAG
         self.split_flag = SPLIT_FLAG
         self.metadata_fields = {
            "xip.reference": "",
            "xip.title": "*",
            "xip.description": "",
            "xip.security_descriptor": "",
            "xip.identifier": "",
            "xip.parent_hierarchy": self.reference,
            "xip.parent_ref": "*",
            "xip.retention_policy_assignment_ref": "*",
            "xip.document_type": "io",
            "xip.retention_policy_assignment_name": self.report_mode,
            "rm.location": "",
            "rm.defaultlocation": "",
            "rm.client": "",
            "rm.coverdates": "",
            "rm.statusdate": ""}

    def report_expired(self, assignments,dict_retention):
        for a in assignments:
            if a.expired: dict_retention.update({'Expired':a.expired})
            else: pass
            
    def report_not_expired(self, assignments,policy,dict_retention):
        for a in assignments:
            if policy.period_unit == "MONTH": p_unit = {'months'}
            else: p_unit = {'years'}
            due_date = datetime.strptime(dict_retention.get('Status_Date'),"%Y-%m-%dT%H:%M:%SZ") + relativedelta(**{p_unit} + int(policy.period))
            dict_retention.update({'Expired': a.expired,'Period':policy.period,'Period_Unit': policy.period_unit,'Due_Date':due_date})

    def report_mode_select(self):
        if self.report_type in {"A","ARCHIVE"}: mode = "ARCHIVE"
        elif self.report_type in {"C","CHECK"}: mode = "CHECK"
        elif self.report_type in {"D","DESTROY"}: mode = "DESTROY"
        elif self.report_type in {"*","ALL"}: mode = "*"
        else: print('Invalid mode, raising SystemExit'); raise SystemExit()
        return mode
    
    def main(self):

        print('Initiating search...')
        list_retention = []
        content.search_callback(content.ReportProgressCallBack())
        report_search = list(content.search_index_filter_list(query="%", filter_values=self.metadata_fields))
        t = len(report_search)
        print(f"Total Results: {str(t)}")
        if not t == 0:
            c = 0
            for f in report_search:
                c += 1
                print(f"Retrieving Expiration Status: {c} / {t}", end="\r")
                asset = entity.asset(f['xip.reference'])
                dict_retention = {'Reference': f.get('xip.reference'),
                                'Document_Type': f.get('xip.document_type'),                               
                                'Title': f.get('xip.title'),
                                'Description': f.get('xip.description'),
                                'Policy_Name': f.get('xip.retention_policy_assignment_name')[0],
                                'Policy_Reference': f.get('xip.retention_policy_assignment_ref')[0],
                                'Box_Location': f.get('rm.location'),
                                'Default_Location': f.get('rm.defaultlocation'),
                                'Client': f.get('rm.client'),
                                'Parent_Ref': f.get('xip.parent_ref'),
                                'Cover_Dates': f.get('rm.coverdates'),
                                'Status_Date': f.get('rm.statusdate')
                                }
                assignments = retention.assignments(asset)
                if len(assignments) > 1: print(f"\n Number of Assignments for {f.get('xip.reference')} is: " + str(len(assignments)))
                if self.expired_flag:
                    self.report_expired(assignments,dict_retention)
                else:
                    policy = retention.policy(reference=f.get('xip.retention_policy_assignment_ref')[0]) 
                    self.report_not_expired(assignments,policy,dict_retention)
                list_retention.append(dict_retention)
            df = pd.DataFrame.from_records(list_retention)
            if self.expired_flag: df = df.dropna(subset=['Expired'])
            if self.path_flag:
                list_path = df[['Reference','Document_Type']].values.tolist()
                list_dict_path = []
                dict_path = {}
                t = len(list_path)
                c = 0
                for record in list_path:
                    c += 1
                    print(f"Retrieving Path: {c} / {t}", end="\r")
                    path,parent_name = path_return(ref=record[0],io=record[1],split_flag=self.split_flag)
                    dict_path.update({'Parent_Name':parent_name,'Path':path})
                    list_dict_path.append(dict_path)
                dfpath = pd.DataFrame.from_records(list_dict_path)
                df.merge(dfpath,left_index=True,right_index=True,how="left")
            boxdf = df.drop(columns=['Reference','Title','Description'])
            boxdf = boxdf.drop_duplicates(subset=['Parent_Ref'])
            boxdf = boxdf.rename({"Parent_Ref":"Reference","Parent_Title":"Title"},axis='columns')
            boxdf['Flag'] = "Box"
            df['Flag'] = "Item"
            df = pd.concat([df,boxdf])
            df.sort_values(by=['Title'])
            if self.report_mode == "*": report_title = "ALL"
            else: report_title = self.report_mode
            output_file = os.path.join(self.output_dir,f'Retention Report_{report_title + "_" + str(datetime.now().strftime("%Y-%m-%d"))}.xlsx')
            with pd.ExcelWriter(output_file,engine='auto',mode='w') as writer:
                df.to_excel(writer,index=False,sheet_name="Master")
                boxdf.to_excel(writer,index=False,sheet_name="Box Summary")
            print(f'File Saved to: {output_file}')
        else: print('No Results Returned...')
        
if __name__ == '__main__':

    startTime = datetime.now()
    print(f"Start time: {str(startTime)}")
    
    ref = "759c51a1-d20c-4b16-b08b-dd3b2719b682"
    output_dir = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Archive\Digital Preservation\github\presRM_tools\presRM_reports"
    RetentionReport(REF=ref,REPORT_TYPE="Archive",PATH_FLAG=True,OUTPUT_DIR=output_dir).main()
    print(f"Complete, running time: {datetime.now() - startTime}")