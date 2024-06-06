from getpass import getpass
from time import sleep
import shlex
import os
try:
    from presRM_labels.labels.labels import LabelGenerator, login_preservica
except Exception as e:
    print(str(e))
    print('Failed to import presRM...')
try:
    import secret
    s_flag = True
except ImportError:
    s_flag = False

def enter_signin(fail_count=0):
    if s_flag:
        print(f'Secret file detected, utilising credentials for: {secret.username}')
        user = secret.username
        passwd = secret.password
    else:
        user = input("""Please enter your Preservica Username: """)
        passwd = getpass("""Please enter your Preservica Password: """)
    try:
        content,entity = login_preservica(username=user,password=passwd)
    except Exception as e:
        if e.args[0] in {401}:    
            print('Error signing in... Please try again...')
            if fail_count == 2: print('This is your last chance...')
            elif fail_count == 3: print('YOU. ARE. DONE.'); sleep(10); raise SystemExit()
            fail_count += 1
            content, entity = enter_signin(fail_count)
        elif e.args[0] in {404,501}:
            print('Preservica is having a moment, please try again later')
        else:
            print('An unknown error occured... Panic!')
            print(e)
            sleep(5)        
    return content,entity

def enter_inputs():
    try:
        boxref = input('Please enter your Box Reference(s) or the location of CSV/XLSX file: ')
        if boxref: boxref = shlex.split(boxref)
        print(boxref)
    except:
        print('Please ensure you enter something here...')
        enter_inputs()
    options = input('If there are any additional options, you want to add please add them here, otherwise click enter\n\nOptions include: [--box-only, --files-only, --combine, --output "path\\to\\output"]: ')
    boxonly = False
    filesonly = False
    combine = False
    toplevel = "de7eaf74-7ad3-4c9b-bab6-ae21008185f0"
    output = os.getcwd()
    if options:
        options = shlex.split(options)
        for count,opt in enumerate(options):
            if "box-only" in opt or "b" == opt: boxonly = True
            if "files-only" in opt or "f" == opt: filesonly = True
            if "output" in opt or "o" == opt:
                output = options[count+1]
                if not os.path.exists(output): print('Output path is invalid and does not exist...'); sleep(5); raise SystemExit()
            if "combine" in opt or "c" == opt: combine = True
            if "toplevel" in opt or "t" == opt: toplevel = options[count+1]
    return boxref,boxonly,filesonly,combine,output,toplevel

def main():
    content,entity = enter_signin()
    boxref,boxonly,filesonly,combine,output,toplevel = enter_inputs()
    print(f"""Search will run with following:
Boxref: {boxref}
Output:{output}
Combine: {combine}
Top-Level: {toplevel}
Box-Only: {boxonly}
Files-Only: {filesonly}
""")
    sleep(2)
    LabelGenerator(content=content,entity=entity,boxref=boxref,output=output,boxonly=boxonly,combine=combine,filesonly=filesonly,toplevel=toplevel).main()

if __name__ == '__main__':
    main()
    print('Complete!')
    input("Please closes this program by banging aimlessly on your keyboard...")
