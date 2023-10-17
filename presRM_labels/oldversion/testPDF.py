from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.barcode import code39,code93,code128
from pyPreservica import *
from lxml import etree 
import argparse

parser = argparse.ArgumentParser(prog="Label Generator",description="Generates Box / File Labels for RM Preservica")
parser.add_argument('boxref',nargs='?')
parser.add_argument('-b','--box-only',action='store_true')
parser.add_argument('-f','--files-only',action='store_true')
parser.add_argument('--top-level',nargs='+',default="de7eaf74-7ad3-4c9b-bab6-ae21008185f0")
args = parser.parse_args()

class LabelGenerator():
     def __init__(self,boxref,boxonly=None,filesonly=None,toplevel=None):
          self.boxref = boxref
          self.box_only = boxonly
          self.files_only = filesonly
          self.top_level = toplevel

content = ContentAPI("chris.prince@unilever.com","Pr3s3rv4t1on",server="unilever.preservica.com")
entity = EntityAPI("chris.prince@unilever.com","Pr3s3rv4t1on",server="unilever.preservica.com")

def add_to_para(string,style):
        para = [Paragraph(string,style)]
        para_if = KeepInFrame(3.5*inch,0.75*inch,para)
        return para_if

def box_label_generation(ref,dept=None,location=None,weight=None,barcode=None):
    if location: location = f"Location: {location}"
    if weight: weight = f"Weight: {weight} kg"
    c = Canvas(f"{fpath}.pdf")
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

def file_label_generation(c,ref,desc=None,location=None,barcode=None):
    c.setPageSize((3*inch,2*inch))
    f1 = Frame(0.25*inch,1.25*inch,2.5*inch,0.5*inch,showBoundary=0)
    f2 = Frame(0.25*inch,0.75*inch,2.5*inch,0.5*inch,showBoundary=0)
    f3 = Frame(0.25*inch,0.25*inch,2.5*inch,0.5*inch,showBoundary=0)
    f1.addFromList([add_to_para(ref,styleSmall)],c)
    f2.addFromList([add_to_para(desc,styleSmall)],c)
    f3.addFromList([add_to_para(location,styleSmall)],c)
    #barcode39.drawOn(c,0.)
    c.showPage()
    return c

if __name__ == "__main__":

    label = LabelGenerator(args.boxref,boxonly=args.box_only,filesonly=args.files_only,toplevel=args.top_level)
    
    #label.boxref = "2020/10" ### To Override / Use in Scripts...

    if not label.boxref: print('Missing Box Reference... Please Provide a Box Reference... Quiting Program...'); raise SystemExit
    if label.files_only and label.box_only: print('Both Box Only and Files Only options are set... Please choose only one...')
    styleNormal = ParagraphStyle('font',fontName="Helvetica",fontSize=20,leading=20)
    styleLarge = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=36)
    styleSmall = ParagraphStyle('font',fontName="Helvetica",fontSize=14,leading=14)

    filters = {"xip.parent_hierarchy": label.top_level,"xip.title":label.boxref,"rm.objtype":"Box","rm.location":"*"}
    def takeFirst(elem):
        return elem[0]
    search = content.search_index_filter_list(query="%",filter_values=filters)
    if not list(search):
        print('No Matches were found in Preservica'); raise SystemExit()
    for hit in dict(search):
        presref = hit.get('xip.reference')
        boxref = hit.get('xip.title')
        boxlocation = str(hit.get('rm.location'))
        dept_lookup = hit.get('xip.parent_hierarchy')[0].split(" ")[-5]
        boxdept = entity.folder(dept_lookup).title
        if label.files_only:
            print('File only selected, skipping File Creation')
            pass
        else:
            fpath = boxref.replace('/','-') + "_Box Label"
            box_label_generation(boxref,dept=boxdept,location=boxlocation,weight="10",barcode=presref)
        if label.box_only:
            print('Box only selected, skipping File Creation')
            pass
        else:
            fpath = boxref.replace('/','-') + "_File Labels"
            # Canvas is declared before loop, to ensure multiple Item PDFs aren't created.
            c = Canvas(f"{fpath}.pdf")
            children = entity.descendants(entity.folder(hit.get('xip.reference')))
            print(list(children))
            for child in children:
                e = entity.entity(child.entity_type,child.reference)
                presref = e.reference
                itemref = e.title
                itemdesc = e.description
                file_label_generation(c,itemref,desc=itemdesc,location=boxlocation,barcode=presref)
            c.save()
            print(fpath)

    print(f'Complete!')