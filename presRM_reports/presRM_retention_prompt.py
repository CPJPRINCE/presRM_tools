from presRM_rentention import login_preservica,RetentionReport
from time import sleep
import os
from getpass import getpass
try: 
    import secret
    s_flag = True
except ImportError:
    s_flag = False

list_of_useful_refs = [["Records Management UK", "bf614a5d-10db-4ff8-addb-2c913b3149eb"],
["Signed Off Records Management UK",	"3cea8f86-1f4e-43dc-81c3-cfa81d5370f0"],
["TM Database", "be7aeb3a-8874-4067-9b62-b46427df8c1d"],
["USCC",	"50de12a0-1466-4d18-9c1a-547450c5579d"],
["Records Management NL", "5f76e749-ddf3-4d3d-a31c-bd07864e9dc2"],
["UK [Root Folder]", "2591783b-725e-4373-9ac1-64e73541e7fa"],
["NL [Root Folder]", "ca45be7a-4a92-4732-8168-4d0f3cef26a2"]]


def raise_invalid():
    print('Invalid Option... Raising SystemExit...')
    sleep(2)
    raise SystemExit()

def enter_signin(fail_count=0):
    if s_flag:
        print(f'Secret file detected, utilising credentials for: {secret.username}')
        user = secret.username
        passwd = secret.password
    else:
        user = input("""Please enter your Preservica Username: """)
        passwd = getpass("""Please enter your Preservica Password: """)
    try:
        entity,retention,content = login_preservica(user=user,passwd=passwd)
    except Exception as e:
        if e.args[0] in {401}:    
            print('Error signing in... Please try again...')
            if fail_count == 2: print('This is your last chance...')
            elif fail_count == 3: print('YOU ARE DONE'); raise SystemExit()
            fail_count += 1
            entity, retention, content = enter_signin(fail_count)
        elif e.args[0] in {404,501}:
            print('Preservica is having a moment, please try again later')
        else:
            print('An unknown error occured... Panic!')
            print(e)
            sleep(5)        
    return entity,retention,content

def enter_inputs():
    print(r"""
Please enter the Preservica Reference of the folder you want to search

*** To use the default settings, simply hit enter, this applies for all options ***
(default is: bf614a5d-10db-4ff8-addb-2c913b3149eb [Records Management UK]

- This searches the selected folder and all subfolders
- You can select also select folder reference on 'lower levels', for limited searches / testing purposes
          
Please see the below for a list of useful references
""")
    headers = ['Folder','Reference']
    print('{:<50s} {:<30s}'.format(*headers))
    for _list in list_of_useful_refs:
        print('{:<50s} {:<30s}'.format(*_list))
    ref_input = input(r"""
                    
Enter reference: """)
    
    if not ref_input:
        ref_input = "bf614a5d-10db-4ff8-addb-2c913b3149eb"
    print(f'Searching folder reference: {ref_input}')

    directory_input = input(rf"""
Please enter the directory, you want to save the report to
                    
- Will default to current working directory, where script is located.
- You can copy and paste this from explorer to make things easier

(default is: {os.getcwd()})

Enter directory: """)
    if not directory_input: directory_input = os.getcwd()
    else: 
        directory_input = directory_input.strip('"')
        if not os.path.exists(directory_input): print('Directory Path doesn\'t exist, please create the directory before using script...'); sleep(10); raise SystemExit()
    print(f'Saving to: {directory_input}')
    report_input = input(r"""
Please enter what report you want to run
            
- The options are ARCHVIVE,CHECK,DESTROY or ALL
- This will return the policies, respectively, assigned to ARCHIVE, CHECK, DESTROY
- ALL will return all policies
- With ALL you can filter ARCHIVE/ in Excel

- Enter: A,C,D or *, to search the respective policy
                        
(default is: *)

Enter A/C/D/*: """)
    report_input = report_input.upper()
    if not report_input: report_input = "*"; print('Using Default: *')
    elif not report_input in {"A","ARCHIVE","C","CHECK","D","DESTROY","*","ALL"}:
        raise_invalid()
    print(f'You selected: {report_input}')

    expired_input = input(r"""
Please select whether you want to run the report on expired items (IE past the retention date) or all items

- Running on all items will include expired items and non-expired items
- For non-expired items a 'Due Date' is also returned, indicating when the retention will expire
- Enter Y for expired items / N for all items

(default is: Y)

Enter Y / N: """)
    expired_input = expired_input.upper()
    if expired_input in {"Y", "YES", True}: expired_input = True; print('Running Expiration Check')
    elif expired_input in {"N", "NO", False}: expired_input = False; print('Not Running Expiration Check')
    elif not expired_input: expired_input = True; print("Using Default: Y")
    else: raise_invalid()

    path_input = input(r"""
Please select whether you want to return the path of the Preservica directoies
    
*** Warning, this significantly increases the time it takes to run the report ***
*** With disabled it takes 1 hour to run (under default options), enabling will take ~3-4 Hours ***
    
- This will return the path delimited by ':::'
- IE Records Management:::UK::Records Management UK:::Top:::Department:::Date:::Box:::Item

(default is: N)

Enter Y / N: """)
    path_input = path_input.upper()
    if path_input in {"Y", "YES", True}:
        print('Running Path Export ***Please be aware this will signifcantly increase runtime, ensure your PC will stay on.')
        path_input = True
        split_input = input(r"""
Please select if you want to remove Prefix and Suffixes from the Path.

- The first three levels IE 'Records Managmenent:::UK:::Records Management UK' will be removed.
- The last level IE ':::Item', ':::Box', ':::Date' will also be removed (depending on level)

(default is: Y)

Enter Y / N: """)
        split_input = split_input.upper()
        if split_input in {"Y", "YES",True}: split_input = True; print('Running prefix and suffix removal on Paths')
        elif split_input in {"N","NO",False}: split_input = False; print('Not running prefix and suffix on Paths')
        elif not split_input: split_input = True; print("Using Default: Y")
        else: print('Invalid option, defaulting to False'); split_input = False 
    elif path_input in {"N","No", False}: path_input = False; split_input = False; ('Not running Path Export')
    elif not path_input: path_input = False; print("Using Default: N"); split_input = False
    else: raise_invalid()
    
    return ref_input,directory_input,report_input,expired_input,path_input,split_input

def main():
    entity,retention,content = enter_signin()
    ref_input,directory_input,report_input,expired_input,path_input,split_input = enter_inputs()    
    RetentionReport(entity=entity,retention=retention,content=content,REF=ref_input,OUTPUT_DIR=directory_input,REPORT_TYPE=report_input,EXPIRED_FLAG=expired_input,PATH_FLAG=path_input,SPLIT_FLAG=split_input).main()
    
if __name__ == "__main__":
    main()
    input("Will close on keyboard mashing...")