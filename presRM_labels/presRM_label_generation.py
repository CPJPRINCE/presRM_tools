from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.barcode import code39,code93,code128
from pyPreservica import *
import argparse
from datetime import datetime
from time import sleep
from PyPDF2 import PdfWriter
try:
    import secret
    s_flag = True
except:
    s_flag = False

def login_preservica(username,password,server="unilever.preservica.com"):
    content = ContentAPI(username,password,server=server)
    entity = EntityAPI(username,password,server=server)
    return content,entity

class LabelGenerator:
    def __init__(self,username=None,password=None,boxref=None,boxonly=False,filesonly=False,toplevel="2591783b-725e-4373-9ac1-64e73541e7fa",output=os.getcwd(),combine=False,barcode_flag=False):
        self.username = username
        self.password = password
        self.boxref = self.check_boxref(boxref)
        self.box_only = boxonly
        self.files_only = filesonly
        self.combine = combine
        self.top_level = toplevel
        self.objtype = "Box"
        self.output_directory = output
        self.content,self.entity = login_preservica(self.username,self.password)        
        self.search_list = self.search_preservica(self.top_level,self.boxref)
        self.barcode_flag = barcode_flag

    def search_preservica(self,top_level,search_ref):
        filters = {"xip.parent_hierarchy": top_level,"xip.title":search_ref,"rm.objtype":self.objtype,"rm.location":"*","rm.weight":"*"}
        search = list(self.content.search_index_filter_list(query="%",filter_values=filters))
        return search
    
    def check_boxref(self,boxref):
        if not boxref: print('Missing Box Reference... Please Provide a Box Reference... Quiting Program...'); raise SystemExit
        for file in boxref:
            if file.endswith('.csv'):
                import csv
                with open(os.path.abspath(file),'r',encoding="utf-8-sig",newline='') as f:
                    read = csv.reader(f)
                    rlist = []
                    for r in read:
                        rlist.append(r[0])
                    boxref = rlist
                    break
            elif file.endswith('.xlsx'):
                import pandas as pd
                df = pd.read_excel(os.path.abspath(file),header=None)
                rlist = df[0].values.tolist()
                boxref = rlist
                break
            else: pass
        return boxref

    def main(self):
        if self.combine: 
            box_combine_list = []
            file_combine_list = []
        if self.files_only and self.box_only: print('Both Box Only and Files Only options are set... Please choose only one...'); raise SystemExit()
        for hit in self.search_list:
            if not list(self.search_list): print(f'\n***\nNo Matches were found in Preservica for: {self.boxref}\n***\n'); continue
            dept_lookup = self.entity.folder(hit.get('xip.parent_hierarchy')[0].split(" ")[-5]).title
            box = BoxLabel(self,location=hit.get('rm.location'),weight=hit.get('rm.weight'),department=dept_lookup)
            box.box_label_generation()
            files = FileLabels(box,self,hit.get('xip.reference'))
            files.file_label_loop()

class BoxLabel(LabelGenerator):
    def __init__(self, LABEL, file_name="_Box Label.pdf", location=None, department=None, weight=None):
        self.boxref = LABEL.boxref
        self.file_name = self.boxref.replace("/","-") + file_name
        self.file_output = os.path.join(LABEL.output_directory,self.file_name)
        self.location = location
        self.department = department
        self.weight = weight
        self.barcode = self.boxref
        self.barcode_flag = LABEL.barcode_flag
        self.styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=18,leading=16)
        self.styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=30)
        self.styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)

    def box_label_generation(self):
        if self.location: self.location = f"Location: {self.location}"
        if self.weight: self.weight = f"Weight: {self.weight} kg"
        C = Canvas(self.file_output)
        C.setPageSize((4*inch,4*inch))
        f1 = Frame(0.25*inch,3.0*inch,3.5*inch,0.75*inch,showBoundary=0)
        f2 = Frame(0.25*inch,2.25*inch,3.5*inch,0.75*inch,showBoundary=0)
        f3 = Frame(0.25*inch,1.5*inch,3.5*inch,0.75*inch,showBoundary=1)
        f4 = Frame(0.25*inch,0.75*inch,3.5*inch,0.75*inch,showBoundary=0)
        if self.barcode_flag:
            bcode = code128.Code128(self.barcode,barHeight=(0.5*inch),humanReadable=True)
            bcode.drawOn(C,0.25*inch,0.25*inch)
        f1.addFromList([add_to_para_box(self.boxref,self.styleLarge)],C)
        f2.addFromList([add_to_para_box(self.department,self.styleNormal)],C)
        f3.addFromList([add_to_para_box(self.location,self.styleNormal)],C)
        f4.addFromList([add_to_para_box(self.weight,self.styleNormal)],C)
        C.save()
        print(f"Box Label Generated to: {self.file_output}")
    
class FileLabels(BoxLabel,LabelGenerator):
    def __init__(self,BOXLABEL,LABEL,xip_reference,file_name="_File Labels.pdf"):
        self.boxref = LABEL.boxref
        self.xipref = xip_reference
        self.file_name = self.boxref.replace('/','-') + file_name 
        self.file_output = os.path.join(LABEL.output_directory,self.file_name)
        self.styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=18,leading=16)
        self.styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=30)
        self.styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)
        self.department = BOXLABEL.department
        self.location = BOXLABEL.location
        self.barcode_flag = LABEL.barcode_flag
        self.entity = LABEL.entity
    
    def file_label_loop(self):
        C = Canvas(self.file_output)
        children = list(self.entity.descendants(self.entity.folder(self.xipref)))
        children.sort(key=lambda x: x.title)
        for child in children:
            e = self.entity.entity(child.entity_type,child.reference)
            self.item_presref = e.reference
            self.item_title = e.title
            self.item_desc = e.description
            self.file_label_generation(C)
        C.save()
        print(f"File Labels Generated to: {self.file_output}")
    
    def file_label_generation(self,C):
        if self.location: location = f"Location: {self.location}"
        C.setPageSize((4*inch,2.25*inch))
        f1 = Frame(0.25*inch,1.7*inch,3.5*inch,0.4*inch,showBoundary=0)
        f2 = Frame(0.25*inch,1.3*inch,2.5*inch,0.4*inch,showBoundary=0)
        f3 = Frame(0.25*inch,0.9*inch,2.5*inch,0.4*inch,showBoundary=0)
        f4 = Frame(0.25*inch,0.55*inch,2.5*inch,0.35*inch,showBoundary=0)
        f1.addFromList([add_to_para_file(self.item_title,self.styleNormal)],C)
        f2.addFromList([add_to_para_file(self.item_desc,self.styleSmall)],C)
        f3.addFromList([add_to_para_file(self.department,self.styleSmall)],C)
        f4.addFromList([add_to_para_file(location,self.styleSmall)],C)
        if self.barcode_flag:
            bcode = code128.Code128(self.item_presref,barHeight=(0.25*inch),barWidth=(0.0080*inch),humanReadable=True)
            bcode.drawOn(C,0.15*inch,0.25*inch)
        C.showPage()
        return C

def add_to_para_box(string,style):
    para = [Paragraph(string,style)]
    para_if = KeepInFrame(3.5*inch,0.75*inch,para)
    return para_if

def add_to_para_file(string,style):
    para = [Paragraph(string,style)]
    para_if = KeepInFrame(2.5*inch,0.3*inch,para)
    return para_if

def parse_args():
    parser = argparse.ArgumentParser(prog="Label Generator",description="Generates Box / File Labels for RM Preservica")
    parser.add_argument('boxref',nargs='+')
    parser.add_argument('-b','--box-only',action='store_true')
    parser.add_argument('-f','--files-only',action='store_true')
    parser.add_argument('--top-level',nargs='?', default="de7eaf74-7ad3-4c9b-bab6-ae21008185f0")
    parser.add_argument('-o', '--output', nargs='?',default=os.getcwd())
    parser.add_argument('-c', '--combine',action='store_true')
    if s_flag:
        try:
            args.username = secret.username
            args.password = secret.password
        except:
            parser.add_argument('-u','--username',nargs='?')
            parser.add_argument('-p','--password',nargs='?')
    args = parser.parse_args()
    return args

if __name__ == "__main__":
    starttime = datetime.now()
    print(starttime)
    args = parse_args()
    LabelGenerator(username=args.username,password=args.password,boxref=args.boxref,boxonly=args.box_only,filesonly=args.files_only,toplevel=args.top_level).main()
    raise SystemExit()

raise SystemExit()
for boxref in args.boxref:
    label = LabelGenerator(boxref,boxonly=args.box_only,filesonly=args.files_only,toplevel=args.top_level,output=args.output)
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
            box_label_generation(boxref,dept=boxdept,location=boxlocation,weight=boxweight,barcode=None)
            if args.combine: box_combine_list.append(fpath)
        if label.box_only:
            print('Box only selected, skipping File Creation')
            pass
        else: 
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