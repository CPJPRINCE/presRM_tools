from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import cm, inch
from reportlab.lib import colors
from reportlab.platypus import Frame,Paragraph,KeepInFrame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


pdf = Canvas("Hello.pdf",pagesize=(4 * inch,4 * inch))

Ref = "2000/100"
Title = "Very long and extranous Department Name"
Loc = "100:101"
Weight = "Weight: 8 kg"
style = getSampleStyleSheet()
Style1 = ParagraphStyle('labeltext1',fontName="Helvetica",fontSize=30,leading=24)

f1 = Frame(0.2*inch,3.1*inch,3*inch,0.80*inch,showBoundary=0)
f2 = Frame(0.2*inch,2.3*inch,3*inch,0.80*inch,showBoundary=0)
f3 = Frame(0.2*inch,1.5*inch,3*inch,0.80*inch,showBoundary=1)
f4 = Frame(0.2*inch,0.7*inch,3*inch,0.80*inch,showBoundary=0)

def add_inframe(text):
    text = [Paragraph(text,Style1)]
    text_if = KeepInFrame(3*inch,1*inch,text)
    return text_if
f1.addFromList([add_inframe(Ref)],pdf)
f2.addFromList([add_inframe(Title)],pdf)
f3.addFromList([add_inframe(Loc)],pdf)
f4.addFromList([add_inframe(Weight)],pdf)
pdf.save()
