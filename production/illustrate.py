"""Illustrate existing adapted DOCX without replacing any original paragraph.
All 39 chapter visuals are native editable Word tables with accessible captions.
8 part-openings and cover are original images. PDF is NOT exported from Word.
"""
from pathlib import Path
from copy import deepcopy
import re,csv,hashlib,zipfile
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from PIL import Image,ImageDraw,ImageFont
from visual_specs import CONCEPTS,PART_THEMES
ROOT=Path(__file__).resolve().parent.parent
SRC=ROOT/'BELIVE_MONEY_manuscrit_final.docx'
OUT=ROOT/'BELIVE_MONEY_manuscrit_illustre.docx'
ASSETS=ROOT/'BELIVE_MONEY_illustrations';ASSETS.mkdir(exist_ok=True)
FONT='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf';BOLD='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
NAVY='#101A35';CYAN='#22C7D5';PURPLE='#7662A6';BLUE='#0B849A'
# Cover is a previously created Arena image, copied as a standalone used file.
import shutil
shutil.copyfile(ROOT/'illustrations/couverture.png',ASSETS/'couverture.png')
PROMPT='Premium editorial book cover illustration only, no text, no letters, no logos. Portrait A4 ratio. Deep midnight navy background, elegant translucent cyan and subtle violet luminous pathways flowing from a human hand holding a simple smartphone at lower left into a constellation of connected circular nodes and upward structured steps, sophisticated geometric vector-meets-3D editorial style, clean strong negative space at top half reserved for title overlay, modern francophone business nonfiction aesthetic, no money symbols, no platform interface.'
# Distinct visual metaphors, manually drawn with Pillow (no generated photo or logos).
def opening(idx):
 w,h=1800,430;im=Image.new('RGB',(w,h),NAVY);d=ImageDraw.Draw(im)
 ox=idx*37
 # eight unique arrangements of one line and connected cards
 pts=[]
 for k in range(5):
  x=210+k*340; y=180+int(55*__import__('math').sin(k*1.25+ox/70))
  if idx%2==0:y=110+(k%3)*90
  if idx%3==0:y=260-(k%3)*85
  pts.append((x,y))
 for a,b in zip(pts,pts[1:]):d.line([a,b],fill=CYAN,width=8)
 for k,(x,y) in enumerate(pts):
  d.rounded_rectangle((x-63,y-56,x+63,y+56),radius=22,fill=('#244D70' if k%2 else '#543E8B'),outline=CYAN,width=5)
  if idx in (0,3,5):d.ellipse((x-16,y-16,x+16,y+16),fill='#F3F7FF')
  elif idx in (1,6):d.line((x-25,y+10,x,y-18,x+25,y+10),fill='white',width=6)
  else:d.rectangle((x-16,y-16,x+16,y+16),outline='white',width=5)
 d.rounded_rectangle((55,46,105,96),radius=14,fill=CYAN)
 # A distinct motif and layout shift for each part, not only a recolored duplicate.
 for m in range(idx+1):
  xx=1340+m*65; yy=65+(idx%3)*12
  d.ellipse((xx-9,yy-9,xx+9,yy+9),fill='#F1F7FF' if m==idx else CYAN)
 d.line((95,365,250+idx*75,365),fill=CYAN,width=5)
 f=ASSETS/f'partie_{idx+1:02d}.png';im.save(f)
 return f
parts=[opening(i) for i in range(8)]
D=Document(SRC)
original=[p.text for p in D.paragraphs];orig_tables=len(D.tables)
existing=[p for p in D.paragraphs if p.style.name=='Heading 1' and re.match(r'^\d{2}\s',p.text)]
assert len(existing)==39
part_paras=[p for p in D.paragraphs if p.style.name=='Partie' and re.match(r'^PARTIE [IVX]+$',p.text)]
assert len(part_paras)==8
# insertion helpers; Word tables are native w:tbl elements (fully editable)
def shading(cell,fill):
 tcPr=cell._tc.get_or_add_tcPr();el=OxmlElement('w:shd');el.set(qn('w:fill'),fill);tcPr.append(el)
def no_break(row):
 trPr=row._tr.get_or_add_trPr();el=OxmlElement('w:cantSplit');trPr.append(el)
def add_visual(chapter,p):
 title,purpose,steps=CONCEPTS[chapter-1]
 # table created at end then moved below chapter heading, preserving all source text
 t=D.add_table(rows=2,cols=len(steps));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
 width=Cm(16/len(steps))
 for i,step in enumerate(steps):
  c=t.cell(0,i);c.width=width;c.text=f'{i+1:02d}'
  shading(c,'101A35');c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
  for r in c.paragraphs[0].runs:r.font.color.rgb=RGBColor(255,255,255);r.bold=True;r.font.size=Pt(9)
  c=t.cell(1,i);c.width=width;c.text=step;shading(c,'EDF5F8' if i%2==0 else 'EDE8F6')
  for pp in c.paragraphs:
   pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
   for r in pp.runs:r.font.size=Pt(9);r.font.color.rgb=RGBColor(16,26,53)
  for row in t.rows:row.cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
 for row in t.rows:no_break(row)
 p._p.addnext(t._tbl)
 # a caption after the table; also preserves visible label for screen-reader reading order
 caption=D.add_paragraph(style='Legende');caption.add_run(f'Figure C{chapter:02d} — {title}. {purpose} Schéma conceptuel éditable, sans données chiffrées.');t._tbl.addnext(caption._p)
 return caption.text
for i,p in enumerate(existing,1):add_visual(i,p)
# Add distinct art after each part title, before chapter start
for i,p in enumerate(part_paras):
 # source part title is the next Heading 1 after this paragraph
 title=p._p.getnext()
 q=D.add_paragraph(style='Legende');q.add_run(f'Ouverture de partie {i+1} — {PART_THEMES[i]}. Illustration conceptuelle originale, sans marque.')
 run=q.add_run();run.add_break();inline=run.add_picture(str(parts[i]),width=Cm(16))
 inline._inline.docPr.set('descr',f'Illustration conceptuelle de la partie {i+1} : cinq étapes reliées, sans logo ni interface.')
 title.addnext(q._p)
# Add alt text to existing images (from previous book) and cover
with zipfile.ZipFile(SRC) as z: source_media=[n for n in z.namelist() if n.startswith('word/media/')]
for p in D.paragraphs:
 for tag in p._p.xpath('.//wp:docPr'):
  if not tag.get('descr'):tag.set('descr','Illustration de couverture ou schéma conceptuel sans logo. Le message figure aussi dans la légende.')
# Make informative table captions explicit; no source text deletion
D.core_properties.title='BELIVE MONEY — édition illustrée adaptée'
D.save(OUT)
# Verify every original text paragraph retains order (tables and captions add but don't delete)
R=Document(OUT); texts=[p.text for p in R.paragraphs]
it=iter(texts)
for para in original:
 assert any(t==para for t in it),f'Original paragraph missing: {para[:75]}'
assert len([p for p in R.paragraphs if p.style.name=='Heading 1' and re.match(r'^\d{2}\s',p.text)])==39
assert len(R.tables)==orig_tables+39
with zipfile.ZipFile(OUT) as z:
 assert z.testzip() is None
 media=[n for n in z.namelist() if n.startswith('word/media/')]
print('Source paragraphs preserved',len(original),'Chapter tables added 39','media total',len(media),'file',OUT.stat().st_size)
# CSV used to compile plan and QA; no irrelevant assets
with open(ROOT/'production/visual_manifest.csv','w',newline='',encoding='utf-8') as f:
 writer=csv.writer(f);writer.writerow(['id','section','objective','type','labels','file','alt'])
 for i,c in enumerate(CONCEPTS,1):writer.writerow([f'C{i:02d}',existing[i-1].text,c[1],'table Word native',' → '.join(c[2]),'internal DOCX',c[1]+': '+', '.join(c[2])])
 for i,fn in enumerate(parts,1):writer.writerow([f'P{i:02d}',f'Partie {i}',PART_THEMES[i-1],'original PNG without text','none',fn.name,f'Five connected steps, part {i}'])
 writer.writerow(['COVER','Couverture','Personne au smartphone et système','AI generated PNG','none','couverture.png','Main tenant un téléphone, réseau lumineux et étapes'])
