"""Independent PDF typesetting from the actual illustrated DOCX XML.
NOT a LibreOffice/Word export; records same text and 39 editable table contents.
"""
from pathlib import Path
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn
from PIL import Image as PIL
from reportlab.platypus import SimpleDocTemplate,Paragraph,PageBreak,Spacer,Image,Table,TableStyle,KeepTogether
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from html import escape
from pypdf import PdfReader
ROOT=Path(__file__).resolve().parent.parent
P=ROOT/'BELIVE_MONEY_manuscrit_illustre.docx';OUT=ROOT/'BELIVE_MONEY_manuscrit_illustre.pdf'
d=Document(P)
pdfmetrics.registerFont(TTFont('DejaVu','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuBold','/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
styles={}
for key,sz,lead,bold,col in [('Normal',9.35,12.6,False,'#101A35'),('Heading 1',19,24,True,'#101A35'),('Heading 2',13,18,True,'#7662A6'),('LivreTitre',24,30,True,'#101A35'),('Partie',27,33,True,'#101A35'),('Etiquette',9,13,True,'#0B849A'),('Encadre',9.5,14,False,'#101A35'),('Prompt',8.8,12.2,False,'#101A35'),('Legende',8.5,12,False,'#7662A6'),('Bibliographie',8.5,12,False,'#101A35')]:
 styles[key]=ParagraphStyle(key,fontName='DejaVuBold' if bold else 'DejaVu',fontSize=sz,leading=lead,textColor=colors.HexColor(col),spaceAfter=7,spaceBefore=8 if key in ('Heading 1','Heading 2','Etiquette') else 0)
body=[];paras=iter(d.paragraphs);tables=iter(d.tables);imgno=0
for elem in d.element.body.iterchildren():
 if elem.tag==qn('w:p'):
  p=next(paras);kind=p.style.name
  if elem.xpath('.//w:br[@w:type="page"]'):body.append(PageBreak())
  if p.text.strip():body.append(Paragraph(escape(p.text).replace('\n','<br/>'),styles.get(kind,styles['Normal'])))
  for blip in elem.xpath('.//a:blip'):
   rid=blip.get(qn('r:embed'))
   if not rid:continue
   data=d.part.related_parts[rid].blob
   im=PIL.open(BytesIO(data));w,h=im.size
   limit=445;maxh=515
   scale=min(limit/w,maxh/h)
   buf=BytesIO(data);buf.seek(0)
   body.append(Image(buf,width=w*scale,height=h*scale));imgno+=1
 elif elem.tag==qn('w:tbl'):
  t=next(tables);data=[]
  for ri,row in enumerate(t.rows):
   cellstyle=ParagraphStyle('cell'+str(ri),parent=styles['Normal'],fontSize=8,leading=11,textColor=colors.white if ri==0 else colors.HexColor('#101A35'))
   data.append([Paragraph(escape(cell.text or ' '),cellstyle) for cell in row.cells])
  if data:
   width=445/len(data[0]);tbl=Table(data,colWidths=[width]*len(data[0]),hAlign='LEFT',repeatRows=0)
   tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#101A35')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('BACKGROUND',(0,1),(-1,-1),colors.HexColor('#EDF5F8')),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#C8DAE3')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
   body.append(tbl);body.append(Spacer(1,9))
def footer(canvas,doc):
 canvas.saveState();canvas.setStrokeColor(colors.HexColor('#CFDFE6'));canvas.line(65,46,530,46)
 canvas.setFont('DejaVu',8);canvas.drawString(65,31,'BELIVE MONEY  ·  Édition illustrée adaptée');canvas.drawRightString(530,31,str(doc.page));canvas.restoreState()
SimpleDocTemplate(str(OUT),pagesize=A4,leftMargin=65,rightMargin=65,topMargin=60,bottomMargin=58,title='BELIVE MONEY — édition illustrée adaptée',author='Valdes Marco BAWA').build(body,onFirstPage=footer,onLaterPages=footer)
r=PdfReader(str(OUT));print('PDF pages',len(r.pages),'images used',imgno,'tables',len(d.tables),'bytes',OUT.stat().st_size)
