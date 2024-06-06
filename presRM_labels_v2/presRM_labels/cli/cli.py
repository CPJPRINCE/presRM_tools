import datetime
import argparse
from presRM_labels.labels.labels import login_preservica, LabelGenerator
try:
    import secret
    s_flag = True
except:
    s_flag = False

def parse_args():
    parser = argparse.ArgumentParser(prog="Label Generator",description="Generates Box / File Labels for RM Preservica")
    parser.add_argument('boxref',nargs='+')
    parser.add_argument('-b','--box-only',action='store_true')
    parser.add_argument('-f','--files-only',action='store_true')
    parser.add_argument('--top-level',nargs='?', default="de7eaf74-7ad3-4c9b-bab6-ae21008185f0")
    parser.add_argument('-o', '--output', nargs='?',default=os.path.dirname(__file__))
    parser.add_argument('-c', '--combine',action='store_true')
    parser.add_argument('-u','--username',nargs='?')
    parser.add_argument('-p','--password',nargs='?')
    args = parser.parse_args()
    return args

if __name__ == "__main__":
    starttime = datetime.now()
    print(starttime)
    args = parse_args()
    if s_flag:
        print(f'Secret file detected, utilising credentials for: {secret.username}')
        user = secret.username
        passwd = secret.password
    else:
        user = args.username
        passwd = args.password    
    content,entity = login_preservica(user,passwd)
    LabelGenerator(content=content,entity=entity,boxref=args.boxref,boxonly=args.box_only,combine=args.combine,filesonly=args.files_only,toplevel=args.top_level,output=args.output).main()
    finishtime = datetime.now() - starttime
    f'Complete! This process took: {finishtime}'