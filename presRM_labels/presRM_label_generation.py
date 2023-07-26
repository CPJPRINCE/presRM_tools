#!"C:\Python311\python.exe"
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.barcode import code39,code93,code128
from pyPreservica import *
import argparse
import sys
from getpass import getpass
from datetime import datetime
from time import sleep
import shlex
from PyPDF2 import PdfWriter

x = input('Please enter your Box Reference(s) or the location of CSV/XLSX file: ')
list("")
if x: sys.argv.extend(shlex.split(x))
y = input('If there are any additional options, you want to add please add them here, otherwise click enter\n \
        Options include: [--box-only, --files-only, --combine, --output "path\\to\\output"] :')
if y: sys.argv.extend(y.split(" "))
u = input('Please enter your Preservica Username (if you are using a secret file, click enter): ')
p = getpass('Please enter your Preservica Password (if you are using a secret file, click enter): ')

parser = argparse.ArgumentParser(prog="Label Generator",description="Generates Box / File Labels for RM Preservica")
parser.add_argument('boxref',nargs='+')
parser.add_argument('-b','--box-only',action='store_true')
parser.add_argument('-f','--files-only',action='store_true')
parser.add_argument('--top-level',nargs='?', default="de7eaf74-7ad3-4c9b-bab6-ae21008185f0")
parser.add_argument('-o', '--output', nargs='?',default=os.getcwd())
parser.add_argument('-c', '--combine',action='store_true')
parser.add_argument('-u','--username',nargs='?')
parser.add_argument('-p','--password',nargs='?')
args = parser.parse_args()
print(args)

if y:
    for yy in y:
        if "box-only" in yy: args.box_only = True
        if "files-only" in yy: args.files_only = True
        if "output" in yy: args.output = yy

if u: args.username = u
if p: args.password = p

starttime = datetime.now()

class LabelGenerator():
     def __init__(self,boxref,boxonly=None,filesonly=None,toplevel=None,output=None):
          self.boxref = boxref
          self.box_only = boxonly
          self.files_only = filesonly
          self.top_level = toplevel
          self.output = output

if not args.username:
    try:
        import secret; args.username = secret.username
    except: print('Please ensure a Username is provided, or a Secret file with Username is provided'); raise SystemExit()
if not args.password:
    try:
        import secret; args.password = secret.password
    except: print('Please ensure a Password is provided, or a Secret file with Password is provided'); raise SystemExit()

content = ContentAPI(args.username,args.password,server="unilever.preservica.com")
entity = EntityAPI(args.username,args.password,server="unilever.preservica.com")

def add_to_para(string,style):
        para = [Paragraph(string,style)]
        para_if = KeepInFrame(3.5*inch,0.75*inch,para)
        return para_if

def add_to_para_file(string,style):
        para = [Paragraph(string,style)]
        para_if = KeepInFrame(2.5*inch,0.3*inch,para)
        return para_if

def box_label_generation(ref,dept=None,location=None,weight=None,barcode=None):
    if location: location = f"Location: {location}"
    if weight: weight = f"Weight: {weight} kg"
    c = Canvas(fpath)
    c.setPageSize((4*inch,4*inch))
    f1 = Frame(0.25*inch,3.0*inch,3.5*inch,0.75*inch,showBoundary=0)
    f2 = Frame(0.25*inch,2.25*inch,3.5*inch,0.75*inch,showBoundary=0)
    f3 = Frame(0.25*inch,1.5*inch,3.5*inch,0.75*inch,showBoundary=1)
    f4 = Frame(0.25*inch,0.75*inch,3.5*inch,0.75*inch,showBoundary=0)
    if barcode:
        bcode = code128.Code128(barcode,barHeight=(0.5*inch),humanReadable=True)
        bcode.drawOn(c,0.25*inch,0.25*inch)
    f1.addFromList([add_to_para(ref,styleLarge)],c)
    f2.addFromList([add_to_para(dept,styleNormal)],c)
    f3.addFromList([add_to_para(location,styleNormal)],c)
    f4.addFromList([add_to_para(weight,styleNormal)],c)
    c.save()
    print(f"Box Label Generated to: {fpath}")

def file_label_generation(c,ref,desc=None,dept=None,location=None,barcode=None):
    if location: location = f"Location: {location}"
    c.setPageSize((4*inch,2.25*inch))
    f1 = Frame(0.25*inch,1.7*inch,3.5*inch,0.4*inch,showBoundary=0)
    f2 = Frame(0.25*inch,1.3*inch,2.5*inch,0.4*inch,showBoundary=0)
    f3 = Frame(0.25*inch,0.9*inch,2.5*inch,0.4*inch,showBoundary=0)
    f4 = Frame(0.25*inch,0.55*inch,2.5*inch,0.35*inch,showBoundary=0)
    f1.addFromList([add_to_para(ref,styleNormal)],c)
    f2.addFromList([add_to_para(desc,styleSmall)],c)
    f3.addFromList([add_to_para(dept,styleSmall)],c)
    f4.addFromList([add_to_para(location,styleSmall)],c)
    if barcode:
        bcode = code128.Code128(barcode,barHeight=(0.25*inch),barWidth=(0.0080*inch),humanReadable=True)
        bcode.drawOn(c,0.15*inch,0.25*inch)
    c.showPage()
    return c

if __name__ == "__main__":
    
    if not args.boxref: print('Missing Box Reference... Please Provide a Box Reference... Quiting Program...'); raise SystemExit

    if args.combine: 
        box_combine_list = []
        file_combine_list = []
    for b in args.boxref:
        if b.endswith('.csv'):
            import csv
            with open(os.path.abspath(b),'r',encoding="utf-8-sig",newline='') as f:
                read = csv.reader(f)
                rlist = []
                for r in read:
                    rlist.append(r[0])
                args.boxref = rlist
                break
        elif b.endswith('.xlsx'):
            import pandas as pd
            df = pd.read_excel(os.path.abspath(b),header=None)
            rlist = df[0].values.tolist()
            args.boxref = rlist
            break
        else: pass

    for boxref in args.boxref:

        label = LabelGenerator(boxref,boxonly=args.box_only,filesonly=args.files_only,toplevel=args.top_level,output=args.output)
        if not label.boxref: print('Missing Box Reference... Please Provide a Box Reference... Quiting Program...'); raise SystemExit
        
        if label.files_only and label.box_only: print('Both Box Only and Files Only options are set... Please choose only one...'); raise SystemExit()
        
        filters = {"xip.parent_hierarchy": label.top_level,"xip.title":label.boxref,"rm.objtype":"Box","rm.location":"*","rm.weight":"*"}
        search = list(content.search_index_filter_list(query="%",filter_values=filters))

        styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=18,leading=16)
        styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=30)
        styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)

        if not list(search): print(f'\n***\nNo Matches were found in Preservica for: {label.boxref}\n***\n'); continue
        for hit in search:
            presref = hit.get('xip.reference')
            boxref = hit.get('xip.title')
            boxweight = hit.get('rm.weight')
            boxlocation = str(hit.get('rm.location'))
            dept_lookup = hit.get('xip.parent_hierarchy')[0].split(" ")[-5]
            boxdept = entity.folder(dept_lookup).title
            if label.files_only:
                print('File only selected, skipping File Creation')
                pass
            else:
                fname = boxref.replace('/','-') + "_Box Label.pdf"
                fpath = os.path.join(label.output,fname)
                box_label_generation(boxref,dept=boxdept,location=boxlocation,weight=boxweight,barcode=presref)
                if args.combine: box_combine_list.append(fpath)
            if label.box_only:
                print('Box only selected, skipping File Creation')
                pass
            else: 
                fname = boxref.replace('/','-') + "_File Labels.pdf"
                fpath = os.path.join(label.output,fname)
                # Canvas is declared before loop, to ensure multiple Item PDFs aren't created.
                c = Canvas(fpath)
                children = list(entity.descendants(entity.folder(hit.get('xip.reference'))))
                children.sort(key=lambda x: x.title)
                for child in children:
                    e = entity.entity(child.entity_type,child.reference)
                    presref = e.reference
                    itemref = e.title
                    itemdesc = e.description
                    file_label_generation(c,itemref,desc=itemdesc,dept=boxdept,location=boxlocation,barcode=presref)
                c.save()
                if args.combine: file_combine_list.append(fpath)
                print(f'File Label Generated to: {fpath}')
    
    if args.combine:
        def combine_pdf(pdf_list,output_path):
            combiner = PdfWriter()
            for pdf in pdf_list:
                if os.path.exists(pdf):
                    combiner.append(pdf)
                try:
                    os.remove(pdf)
                except Exception as e:
                    print(f'Error: {e}')
                    print('Possibly Multiple Matches in Preservica Search ')
            combiner.write(output_path)
            print(f'Combined PDFs to: {output_path}')
            combiner.close()
        print('Combining PDFs...')
        if box_combine_list:
            combine_pdf(box_combine_list,os.path.join(label.output,"Combined_Box Labels.pdf"))
        if file_combine_list:
            combine_pdf(file_combine_list,os.path.join(label.output,"Combined_File Labels.pdf"))
        

    finishtime = datetime.now() - starttime
    print(f'Complete! This process took: {finishtime}')
    sleep(5)
    print('Now go print them labels!!')
    sleep(2)