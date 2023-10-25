from getpass import getpass
import time
import shlex
import sys
try:
    from presRM_label_generation import LabelGenerator,login_preservica
except:
    print('Failed to import presRM...')
try:
    import secret
    s_flag = True
except ImportError:
    s_flag = False

print("Goodbe")
def enter_signin(fail_count=0):
    if s_flag:
        print(f'Secret file detected, utilising credentials for: {secret.username}')
        user = secret.username
        passwd = secret.password
    else:
        user = input("""Please enter your Preservica Username: """)
        passwd = getpass("""Please enter your Preservica Password: """)
    try:
        entity,content = login_preservica(user=user,passwd=passwd)
    except Exception as e:
        if e.args[0] in {401}:    
            print('Error signing in... Please try again...')
            if fail_count == 2: print('This is your last chance...')
            elif fail_count == 3: print('YOU ARE DONE'); raise SystemExit()
            fail_count += 1
            entity, content = enter_signin(fail_count)
        elif e.args[0] in {404,501}:
            print('Preservica is having a moment, please try again later')
        else:
            print('An unknown error occured... Panic!')
            print(e)
            time.sleep(5)        
    return entity,content,user,passwd

def enter_inputs():
    try:
        boxref = input('Please enter your Box Reference(s) or the location of CSV/XLSX file: ')
        if boxref: sys.argv.extend(shlex.split(boxref))
    except:
        print('Please ensure you enter something here...')
        enter_inputs()
    options = input('If there are any additional options, you want to add please add them here, otherwise click enter\n \
            Options include: [--box-only, --files-only, --combine, --output "path\\to\\output"] :')
    if options:
        options = sys.argv.extend(shlex.split(options))
        for count,opt in enumerate(options):
            if "box-only" in opt: box_only = True
            else: box_only = False
            if "files-only" in opt: files_only = True
            else: files_only = False
            if "output" in opt: output = options[count +1]
            if "combine" in opt: combine = True
            else: combine = False
            if "toplevel" in opt: toplevel = options[count +1]
            else: toplevel = "de7eaf74-7ad3-4c9b-bab6-ae21008185f0"
    return boxref,box_only,files_only,combine,output,toplevel

def main():
    entity,content,username,password = enter_signin()
    username,password,boxref,box_only,files_only,combine,toplevel = enter_inputs()
    LabelGenerator(username=username,password=password,boxref=boxref,box_only=box_only,combine=combine,files_only=files_only,top_level=toplevel).entity,content = entity,content

if __name__ == '__main__':
    main()
    input("Will close on keyboard mashing...") 