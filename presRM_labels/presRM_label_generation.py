from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode import code128
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
    if username == None or password == None:
        print('Username or Password is set to none... Please address...')
        raise SystemExit()
    content = ContentAPI(username,password,server=server)
    entity = EntityAPI(username,password,server=server)
    print('Successfully logged into Preservica')
    return content,entity

class LabelGenerator:
    def __init__(self,content,entity,boxref=None,boxonly=False,filesonly=False,toplevel="2591783b-725e-4373-9ac1-64e73541e7fa",output=os.path.dirname(__file__),combine=False,barcode_flag=False):
        self.content = content
        self.entity = entity
        self.box_list = self.check_boxref(boxref)
        self.box_only = boxonly
        self.files_only = filesonly
        self.combine = combine
        self.top_level = toplevel
        self.objtype = "Box"
        self.output_directory = output
        self.barcode_flag = barcode_flag
                
    def search_preservica(self,top_level,search_ref):
        filters = {"xip.parent_hierarchy": top_level,"xip.title":search_ref,"rm.objtype":self.objtype,"rm.location":"","rm.weight":""}
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
                boxref = df[0].values.tolist()
                break
            else: pass
        return boxref
    
    def combine_pdf(self,pdf_list,output_path):
        print(f'Combined PDFs to: {output_path}')
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
        combiner.close()

    def main(self):
        os.chdir(os.path.dirname(__file__))
        if self.combine: 
            box_combine_list = []
            file_combine_list = []
        if self.files_only and self.box_only: print('Both Box Only and Files Only options are set... Please choose only one...'); raise SystemExit()
        for box in self.box_list:
            box_label = BoxLabel(self,box)
            if box_label.match_flag:
                if not self.files_only:
                    box_output = box_label.box_label_generation()
                    if self.combine: box_combine_list.append(box_output)
                if not self.box_only:
                    files = FileLabels(box_label,self)
                    file_output = files.file_label_loop()
                    if self.combine: file_combine_list.append(file_output)
            else: pass
        if self.combine:
            if box_combine_list:
                self.combine_pdf(box_combine_list,os.path.join(self.output_directory,"Combined_Box Labels.pdf"))
            if file_combine_list:
                self.combine_pdf(file_combine_list,os.path.join(self.output_directory,"Combined_File Labels.pdf"))

class BoxLabel(LabelGenerator):
    def __init__(self, LABEL, box, file_name="_Box Label.pdf"):
        self.entity = LABEL.entity
        self.barcode_flag = LABEL.barcode_flag
        self.combine = LABEL.combine
        self.boxref = box
        self.file_name = self.boxref.replace("/","-") + file_name
        self.file_output = os.path.join(LABEL.output_directory,self.file_name)
        self.styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=24,leading=24,valign='middle')
        self.styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=34)
        self.styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)
        self.box_search = LABEL.search_preservica(LABEL.top_level,self.boxref)
        self.uarmimage = os.path.relpath("./uarm.jpg")
        self.match_flag = self.match_check() 
        self.box_label_values()
        
    def match_check(self):
        if not list(self.box_search):print(f'\n***\nNo Matches were found in Preservica for: {self.boxref}\n***\n');return False
        else: return True
    
    def box_label_values(self):
        for x in self.box_search:
            self.xipref = x.get('xip.reference')
            self.boxtitle = x.get('xip.title')
            self.location = x.get('rm.location')
            self.weight = x.get('rm.weight')
            dept_lookup = x.get('xip.parent_hierarchy')[0].split(" ")[-5]
            self.department = self.entity.folder(dept_lookup).title
    
    def box_label_generation(self):
        if self.location: location = f"Location: {self.location}"
        if self.weight: weight = f"Weight: {self.weight} kg"
        else: weight = ""
        C = Canvas(self.file_output)
        C.setPageSize((4*inch,4*inch))
        f1 = Frame(0.2*inch,2.9*inch,2.8*inch,0.9*inch,showBoundary=0)
        f2 = Frame(0.2*inch,1.1*inch,3.6*inch,1.8*inch,showBoundary=0)
        f3 = Frame(0.2*inch,0.65*inch,1.8*inch,0.45*inch,showBoundary=1)
        f4 = Frame(0.2*inch,0.2*inch,1.8*inch,0.45*inch,showBoundary=1)
        if self.barcode_flag:
            bcode = code128.Code128(self.barcode,barHeight=(0.5*inch),humanReadable=True)
            bcode.drawOn(C,0.25*inch,0.25*inch)
        f1.addFromList([add_to_para_box(self.boxtitle,self.styleLarge)],C)
        f2.addFromList([add_to_para_box(self.department,self.styleNormal)],C)
        f3.addFromList([add_to_para_box(location,self.styleNormal)],C)
        if weight: f4.addFromList([add_to_para_box(weight,self.styleNormal)],C)
        try: C.drawInlineImage(self.uarmimage,2.8*inch,0.4*inch,width=1.0*inch,preserveAspectRatio=True)
        except: pass
        C.save()
        print(f"Box Label Generated to: {self.file_output}")
        return self.file_output
        
class FileLabels(BoxLabel,LabelGenerator):
    def __init__(self,BOXLABEL,LABEL,file_name="_File Labels.pdf"):
        self.file_name = BOXLABEL.boxref.replace('/','-') + file_name 
        self.file_output = os.path.join(LABEL.output_directory,self.file_name)
        self.styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=20,leading=18)
        self.styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=28,leading=26)
        self.styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)
        self.entity = LABEL.entity
        self.barcode_flag = LABEL.barcode_flag
        self.xipref = BOXLABEL.xipref
        self.location = BOXLABEL.location
        self.department = BOXLABEL.department
        self.uarmimage = BOXLABEL.uarmimage
        
    def file_label_loop(self):
        C = Canvas(self.file_output)
        children = list(self.entity.descendants(self.entity.folder(self.xipref)))
        children.sort(key=lambda x: "-".join([str(x).zfill(5) for x in str(x.title).split("/")]))
        #children.sort(key=lambda x: [str(x.title).split()])
        for child in children:
            e = self.entity.entity(child.entity_type,child.reference)
            self.item_presref = e.reference
            self.item_title = e.title
            self.item_desc = e.description
            self.file_label_generation(C)
        C.save()
        print(f"File Labels Generated to: {self.file_output}")
        return self.file_output
    
    def file_label_generation(self,C):
        if self.location: location = f"Location: {self.location}"
        C.setPageSize((4*inch,2.25*inch))
        f1 = Frame(0.25*inch,1.625*inch,3.0*inch,0.5*inch,showBoundary=0)
        f2 = Frame(0.25*inch,1.125*inch,3.5*inch,0.5*inch,showBoundary=0)
        f3 = Frame(0.25*inch,0.625*inch,3.5*inch,0.5*inch,showBoundary=0)
        f4 = Frame(0.25*inch,0.125*inch,3.5*inch,0.5*inch,showBoundary=0)
        f1.addFromList([add_to_para_file(self.item_title,self.styleNormal)],C)
        f2.addFromList([add_to_para_file(self.item_desc,self.styleSmall)],C)
        f3.addFromList([add_to_para_file(self.department,self.styleSmall)],C)
        f4.addFromList([add_to_para_file(location,self.styleSmall)],C)
        if self.barcode_flag:
            bcode = code128.Code128(self.item_presref,barHeight=(0.25*inch),barWidth=(0.0080*inch),humanReadable=True)
            bcode.drawOn(C,0.15*inch,0.25*inch)
        try: C.drawInlineImage(self.uarmimage,3.0*inch,-1.1*inch,width=0.75*inch,preserveAspectRatio=True) 
        except: pass            
        C.showPage()
        return C

def add_to_para_box(string,style):
    para = [Paragraph(string,style)]
    para_if = KeepInFrame(3.5*inch,0.9*inch,para)
    return para_if

def add_to_para_file(string,style):
    para = [Paragraph(string,style)]
    para_if = KeepInFrame(3.5*inch,0.5*inch,para)
    return para_if

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