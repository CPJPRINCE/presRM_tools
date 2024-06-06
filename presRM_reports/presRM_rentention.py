from pyPreservica import *
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
try:
    import secret
except:
    pass
import numpy as np

def login_preservica(user,passwd,sever="unilever.preservica.com"):
    entity = EntityAPI(username=user,password=passwd, server=sever)
    retention = RetentionAPI(username=user,password=passwd, server=sever)        
    content = ContentAPI(username=user,password=passwd, server="unilever.preservica.com")
    print('Successfully logged into Preservica')
    return entity,retention,content

class RetentionReport():
    def __init__(self,entity=None, content=None, retention=None, REF="bf614a5d-10db-4ff8-addb-2c913b3149eb",OUTPUT_DIR=os.getcwd(),REPORT_TYPE="ALL",EXPIRED_FLAG=True,PATH_FLAG=False,SPLIT_FLAG=True):
         self.reference = REF 
         self.output_dir = os.path.abspath(OUTPUT_DIR)
         self.report_type = REPORT_TYPE.upper()
         self.report_mode = self.report_mode_select()
         self.expired_flag = EXPIRED_FLAG
         self.path_flag = PATH_FLAG
         self.split_flag = SPLIT_FLAG
         self.entity = entity
         self.content = content
         self.retention = retention
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
            "rm.coverdate": "",
            "rm.statusdate": "1900-01-01 - 3000-01-01"}
    def parent_return(self,ref,io):
        if io == "IO": x = self.entity.asset(ref)
        elif io == "SO": x = self.entity.folder(ref)
        else: print("Error?")
        while x.parent is not None:
            f = self.entity.folder(x.parent)
            parent_name = f.title
            parent_desc = f.description
        return parent_name,parent_desc

    def path_return(self,ref,io,split_flag):
        try:
            if io == "IO": x = self.entity.asset(ref)
            elif io == "SO": x = self.entity.folder(ref)
            else: print("Error?")
            fpath = x.title
            f = self.entity.folder(x.parent)
            parent_name = f.title    
            parent_desc = f.description
            while f.parent is not None:        
                fpath = f.title + ":::" + fpath
                f = self.entity.folder(f.parent)    
            if split_flag:
                fpath = fpath.rsplit(":::",maxsplit=1)[0]
                fpath = fpath.split(":::",maxsplit=3)[-1]
        except Exception as e:
            print(e)
            fpath = "ERROR"
            parent_name = "ERROR"
            parent_desc = "ERROR"
        return fpath,parent_name,parent_desc
    
    def report_expired(self, assignments,policy,dict_retention):
        for a in assignments:
            if a.expired:
                if policy.period_unit == "MONTH": p_unit = {'months':int(policy.period)}
                else: p_unit = {'years':int(policy.period)}
                due_date = dict_retention.get('Status_Date') + relativedelta(**p_unit)
                dict_retention.update({'Expired': a.expired,'Due_Date':due_date})            
            else: pass
            
    def report_not_expired(self, assignments,policy,dict_retention):
        for a in assignments:
            if policy.period_unit == "MONTH": p_unit = {'months':int(policy.period)}
            else: p_unit = {'years':int(policy.period)}
            due_date = dict_retention.get('Status_Date') + relativedelta(**p_unit)
            dict_retention.update({'Expired': a.expired,'Period':policy.period,'Period_Unit': policy.period_unit,'Due_Date':due_date})

    def report_mode_select(self):
        if self.report_type in {"A","ARCHIVE"}: mode = "ARCHIVE"
        elif self.report_type in {"C","CHECK"}: mode = "CHECK"
        elif self.report_type in {"D","DESTROY"}: mode = "DESTROY"
        elif self.report_type in {"*","ALL"}: mode = "*"
        else: print('Invalid mode, raising SystemExit'); raise SystemExit()
        return mode
    
    def export_xl(self,df,boxdf,output_file):
        try:
            with pd.ExcelWriter(output_file,engine='auto',mode='w') as writer:
                df.to_excel(writer,index=False,sheet_name="Master")
                boxdf.to_excel(writer,index=False,sheet_name="Box Summary")
            print(f'File Saved to: {output_file}')
        except Exception as e:
            print(e)
            print(f'Save failed... Please ensure file is closed: {output_file}')
            time.sleep(10)
            self.export_xl(df,boxdf,output_file)
    
    def main(self):
        print('Initiating search...')
        list_retention = []
        self.content.search_callback(self.content.ReportProgressCallBack())
        report_search = list(self.content.search_index_filter_list(query="%", filter_values=self.metadata_fields,page_size=1000))
        t = len(report_search)
        print()
        print(f"Total Results: {str(t)}")
        if not t == 0:
            c = 0
            for f in report_search:
                c += 1
                print(f"Retrieving Expiration Status: {c} / {t}", end="\r")
                reference = f.get('xip.reference')
                asset = self.entity.asset(f.get('xip.reference'))
                sdate_temp = f.get('rm.statusdate')
                if sdate_temp:
                    try:
                        if sdate_temp >0:
                            sdate_temp = datetime.fromtimestamp(sdate_temp/1000)
                        else:
                            sdate_temp = datetime(1970, 1, 1) + timedelta(seconds=int(sdate_temp/1000))
                    except Exception as e:
                        print(f'\nInvalid datetime in reference: {reference}; this record will be skipped\n')
                        print(e)
                        continue
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
                                'Cover_Dates': f.get('rm.coverdate'),
                                'Status_Date': sdate_temp                                
                                }
                try: 
                    assignments = self.retention.assignments(asset)
                except Exception as e: 
                    print(f'There was an error retrieving Retention Assignment for: {reference}')
                    print(e)
                    continue
                if len(assignments) > 1: print(f"\n Number of Assignments for {f.get('xip.reference')} is: " + str(len(assignments)))
                try:
                    policy = self.retention.policy(reference=f.get('xip.retention_policy_assignment_ref')[0])
                except Exception as e:
                    print(f'There was an error retrieving Retention Policy for: {reference}')
                    print(e)
                    continue                    
                if self.expired_flag:
                    self.report_expired(assignments,policy,dict_retention)
                else:
                    self.report_not_expired(assignments,policy,dict_retention)
                list_retention.append(dict_retention)
            print('Collecting Expiration Status Finished')
            df = pd.DataFrame.from_records(list_retention)
            print('Records Successfully loaded into DataFrame')
            print('\n')
            if self.expired_flag: df = df.dropna(subset=['Expired'])
            if self.path_flag:
                list_path = df[['Reference','Document_Type']].values.tolist()
                list_dict_path = []
                t = len(list_path)
                c = 0
                for r in list_path:
                    dict_path = {}
                    c += 1
                    print(f"Retrieving Path: {c} / {t}", end="\r")
                    path,parent_name,parent_desc = self.path_return(ref=r[0],io=r[1],split_flag=self.split_flag)
                    dict_path.update({'Reference':r[0],'Parent_Name':parent_name,'Parent_Desc': parent_desc,'Path':path})
                    list_dict_path.append(dict_path)
                dfpath = pd.DataFrame.from_records(list_dict_path)
                df = df.merge(dfpath,on="Reference",how="left")
            print()
            boxdf = df.drop_duplicates(subset=['Parent_Ref'])
            if not self.path_flag:
                boxdf.reset_index(drop=True)
                boxdf['Parent_Name'] = boxdf['Title'].str.rsplit("/",n=1).str[0]
            else:
                boxdf['Path'] = boxdf['Path'].str.rsplit(":::",n=1).str[0]
            boxdf = boxdf.drop(columns=['Reference','Title','Description','Cover_Dates'])                
            boxdf = boxdf.rename({"Parent_Ref":"Reference","Parent_Name":"Title","Parent_Desc":"Description"},axis='columns')                    
            boxdf['Flag'] = "Box"
            df['Flag'] = "Item"
            df = pd.concat([df,boxdf],ignore_index=True)

            #df['SortHelper'] = df['Title'].str.split('/',expand=True).fillna(value="0")
            df['SortHelper'] = df['Title'].apply(lambda x: "-".join([str(x).zfill(5) for x in str(x).split("/")]))
            df = df.sort_values(by=['SortHelper'])
            #df = df.drop(columns='SortHelper')
            if self.report_mode == "*": report_title = "ALL"
            else: report_title = self.report_mode
            output_file = os.path.join(self.output_dir,f'Retention Report_{self.reference + "_" + report_title + "_" + str(datetime.now().strftime("%Y-%m-%d"))}.xlsx')
            self.export_xl(df,boxdf,output_file)
        else: print('No Results Returned...')
        
if __name__ == '__main__':
    startTime = datetime.now()
    print(f"Start time: {str(startTime)}")
    entity,retention,content = login_preservica(user=secret.username,passwd=secret.password)
    ref = "2591783b-725e-4373-9ac1-64e73541e7fa"
    #output_dir = r"C:\Users\Chris.Prince\Unilever\UARM Teamsite - Archive\Digital Preservation\github\presRM_tools\presRM_reports"
    try:
        RetentionReport(entity=entity,retention=retention,content=content,REF=ref,REPORT_TYPE="All",EXPIRED_FLAG=True,PATH_FLAG=False).main()
        print(f"Complete, running time: {datetime.now() - startTime}")
    except Exception as e:
        print('Process Failed...')
        time.sleep(2)
        print('Reason for Failure: ')
        print(e)
        input('\n Will close on pressing enter...')
        raise SystemError()
