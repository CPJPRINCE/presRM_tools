import sys
sys.path.append("")

from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.graphics.barcode import code128
from PyPDF2 import PdfWriter
import os
import pandas as pd
import csv
from presRM_labels.labels.options import *
from presRM_labels.labels.preservica import search_preservica,login_preservica,CONTENT,ENTITY

class LabelGenerator:
    def __init__(self,search_term: list, top_level: str ="2591783b-725e-4373-9ac1-64e73541e7fa",
                 box_only: bool = False,
                 files_only: bool =False,
                 output: os.path = os.path.dirname(__file__),
                 combine_flag: bool = False,
                 barcode_flag: bool = False):
        self.content = CONTENT
        self.entity = ENTITY
        self.search_list = self._return_search_list(search_term)
        self.box_only = box_only
        self.files_only = files_only
        self.files_mode = "descendants"
        self.combine_flag = combine_flag
        self.combine_list = []
        self.top_level = top_level
        self.output_directory = output
        self.barcode_flag = barcode_flag
                
    def _return_search_list(self,search_term):
        if not search_term: print('Missing search term... Please provide a search term... Quiting Program...'); raise SystemExit
        for file in search_term:
            if file.endswith('.csv'):
                with open(os.path.abspath(file),'r',encoding="utf-8-sig",newline='') as f:
                    read = csv.reader(f)
                    rowlist = []
                    for row in read: rowlist.append(row[0])
                    search_list = rowlist
                    break
            elif file.endswith('.xlsx'):
                df = pd.read_excel(os.path.abspath(file),header=None)
                search_list = df[0].values.tolist()
                break
            else: pass
        return search_list
    
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

    def generate_box_label(self):
        for term in self.search_list:
            box_label = Label(term,self.top_level,frames=BOX_FRAMES,logo=LOGO,page_size=BOX_PAGE_SIZE)
            if self.combine_flag: self.combine_list.append(box_label.file_path)
        if self.combine_flag: self.combine_pdf(self.combine_list)
        return box_label.search

    def generate_file_label_descendents():
        pass

    def generate_file_label_content():
        pass
    
    def generate_file_label_single():
        pass

    def main(self):
        os.chdir(os.path.dirname(__file__))
        if self.combine: 
            box_combine_list = []
            file_combine_list = []
        if self.files_only and self.box_only: print('Both Box Only and Files Only options are set... Please choose only one...'); raise SystemExit()
        if not self.files_only:
            self.generate_box_label()
        if not self.box_only:
            if self.files_mode == "descendants":
                self.generate_file_label_descendents()
            elif self.files_mode == "content":
                self.generate_file_label_content()
            elif self.files_mode == "exact":
                self.generate_file_label_single()
            else: pass
        if self.combine:
            if box_combine_list:
                self.combine_pdf(box_combine_list,os.path.join(self.output_directory,"Combined_Box Labels.pdf"))
            if file_combine_list:
                self.combine_pdf(file_combine_list,os.path.join(self.output_directory,"Combined_File Labels.pdf"))

class Label():
    def __init__(self, search_term: str = None,
                 top_level: str = None,
                 mode: set = None,
                 frames: list = None,
                 logo: str = None,
                 page_size: str = None,
                 output_directory: os = os.getcwd(), 
                 file_name: str = "_Label.pdf"):
        
        self.search_term = search_term
        self.mode_flag = mode
        self.output_path = os.path.join(output_directory,self.search_term + file_name)
        self.logo = logo
        self.frames = frames
        self.page_size = page_size
        self.canvas = Canvas(self.output_path)
        if self.mode_flag == "desdendants":
            pass
        else:
            self.search = search_preservica()
            self.match_flag = self.match_check()
            if self.match_flag is not None: 
                if self.match_flag == "Multi":
                    for s in self.search:
                        self.dict_update(s)
                        self.label_generation()
                        self.canvas.showPage()
                    self.canvas.save()
                elif self.match_flag == "Single":
                    for s in self.search:
                        self.dict_update(s)
                    self.label_generation()
                    self.canvas.save()
            else: pass
    
    def dict_update(self,search_dict):
        for key in search_dict.keys():
            for f in self.frames:
                if key == f.get('Field'):
                    f.update({'Value':search_dict.get(key)})

    def match_check(self):
        if not list(self.search):print(f'\n***\nNo Matches were found in Preservica for: {self.search_term}\n***\n'); return None
        elif len(list(self.search)) > 1:print(f'There are multiple matches to this box. Will print all.'); return "Multi"
        else: return "Single"
        
        # for x in self.box_search:
        #     self.xipref = x.get('xip.reference')
        #     self.boxtitle = x.get('xip.title')
        #     self.location = x.get('rm.location')
        #     self.weight = x.get('rm.weight')
        #     dept_lookup = x.get('xip.parent_hierarchy')[0].split(" ")[-5]
        #     self.department = self.entity.folder(dept_lookup).title
    
    def label_generation(self):
        self.canvas.setPageSize(self.page_size)
        for frame in self.frames:
            f = frame.get('Frame')
            text = frame.get('Field')
            if frame.get('Additional Text') is not None: text = frame.get('Additional Text') + text
            style = frame.get('Style')
            f.addFromList([add_to_para_box(text,style)],self.canvas)
        # if self.barcode_flag:
        #     bcode = code128.Code128(self.barcode,barHeight=(0.5*inch),humanReadable=True)
        #     bcode.drawOn(C,0.25*inch,0.25*inch)
        # try: C.drawInlineImage(self.logo,2.8*inch,0.4*inch,width=1.0*inch,preserveAspectRatio=True)
        # except: pass
        
class FileLabels(BoxLabel,LabelGenerator):
    def __init__(self,BOXLABEL,LABEL,file_name="_File Labels.pdf"):
        self.file_name = BOXLABEL.boxref.replace('/','-') + file_name 
        self.file_output = os.path.join(LABEL.output_directory,self.file_name)
        self.entity = LABEL.entity
        self.barcode_flag = LABEL.barcode_flag
        self.xipref = BOXLABEL.xipref
        self.location = BOXLABEL.location
        self.department = BOXLABEL.department
        self.logo = BOXLABEL.logo
        
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
        C.setPageSize(FILE_PAGE_SIZE)
        for n, F in enumerate(FILE_FRAMES):
            F.addFromList([add_to_para_file()],C)

        f1.addFromList([add_to_para_file(self.item_title,self.styleNormal)],C)
        f2.addFromList([add_to_para_file(self.item_desc,self.styleSmall)],C)
        f3.addFromList([add_to_para_file(self.department,self.styleSmall)],C)
        f4.addFromList([add_to_para_file(location,self.styleSmall)],C)
        if self.barcode_flag:
            bcode = code128.Code128(self.item_presref,barHeight=(0.25*inch),barWidth=(0.0080*inch),humanReadable=True)
            bcode.drawOn(C,0.15*inch,0.25*inch)
        try: C.drawInlineImage(self.logo,3.0*inch,-1.1*inch,width=0.75*inch,preserveAspectRatio=True) 
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