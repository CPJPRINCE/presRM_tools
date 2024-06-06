from reportlab.lib.units import inch
from reportlab.platypus import Paragraph,KeepInFrame, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.graphics.barcode import code128
import os.path

top_level = ""
search_ref = ""
objtype = ""


BOX_FIELDS = {"xip.title","rm.location","rm.weight"}
FILE_FIELDS = {}

BOX_PAGE_SIZE = (4*inch,4*inch)
FILE_PAGE_SIZE = (4*inch,2.25*inch)

BOX_FRAMES = [Frame(0.2*inch,2.9*inch,2.8*inch,0.9*inch,showBoundary=0),
        Frame(0.2*inch,1.1*inch,3.6*inch,1.8*inch,showBoundary=0),
        Frame(0.2*inch,0.65*inch,1.8*inch,0.45*inch,showBoundary=1),
        Frame(0.2*inch,0.2*inch,1.8*inch,0.45*inch,showBoundary=1)]

FILE_FRAMES = [Frame(0.25*inch,1.625*inch,3.0*inch,0.5*inch,showBoundary=0),
        Frame(0.25*inch,1.125*inch,3.5*inch,0.5*inch,showBoundary=0),
        Frame(0.25*inch,0.625*inch,3.5*inch,0.5*inch,showBoundary=0),
        Frame(0.25*inch,0.125*inch,3.5*inch,0.5*inch,showBoundary=0)]

STYLE_NORMAL = ParagraphStyle('font',fontName="Helvetica",fontSize=24,leading=24,valign='middle')
STYLE_LARGE = ParagraphStyle('font',fontName="Helvetica",fontSize=36,leading=34)
STYLE_SMALL = ParagraphStyle('font',fontName="Helvetica",fontSize=16,leading=14)

STYLE_DICT = {"xip.title":STYLE_NORMAL,"rm.location":STYLE_SMALL,"rm.weight":STYLE_SMALL}
ADD_TEXT_DICT = {"rm.location": "Location: ","rm.weight":"Weight (kg): "}

LABEL_ITEMS = []
for FIELD, FRAME in BOX_FIELDS,BOX_FRAMES:
    COMPILED_FRAME = {"Frame":FRAME,"Field":FIELD}
    COMPILED_FRAME.update({"Style":STYLE_DICT.get(FIELD),"Additional Text":ADD_TEXT_DICT.get(FIELD)})
    LABEL_ITEMS.append(COMPILED_FRAME)

LOGO = os.path.relpath("../assets/logo.jpg")
