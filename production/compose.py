"""Compose a reader edition from the supplied conversational text adaptation.
No original DOCX exists. The PDF is separately typeset, not a converted rendering of the DOCX.
"""
import csv,re,zipfile,io,html,os
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER,TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image as PILImage,ImageDraw,ImageFont
from pypdf import PdfReader
from contenu import CHAPTERS,PARTS,PRELIMS,TITLE,SUBTITLE,AUTHOR
from annexes import ANNEX_PROMPTS,GLOSSARY,CALENDAR,SOURCES
ROOT=Path(__file__).resolve().parent.parent
NAVY='#101a35';CYAN='#008ea3';PURPLE='#7854af';LIGHT='#edf5f8'
# Convert our SVG concepts into crisp PNG via Pillow, using text as accessible native book captions.
fontpath='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
boldpath='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
def card(path,labels):
 im=PILImage.new('RGB',(1800,450),'#f4f8fc'); d=ImageDraw.Draw(im); f=ImageFont.truetype(boldpath,27)
 w=(1680-25*(len(labels)-1))//len(labels)
 for i,label in enumerate(labels):
  x=60+i*(w+25); d.rounded_rectangle((x,145,x+w,290),radius=26,fill=NAVY if i%2==0 else PURPLE)
  words=label.split(); lines=[''];
  for word in words:
   if d.textbbox((0,0),lines[-1]+' '+word,font=f)[2]>w-38 and lines[-1]:lines.append('')
   lines[-1]=(lines[-1]+' '+word).strip()
  for n,line in enumerate(lines[:2]):
   bb=d.textbbox((0,0),line,font=f);d.text((x+w/2-(bb[2]-bb[0])/2,191+(n-(len(lines[:2])-1)/2)*32),line,font=f,fill='white')
  if i<len(labels)-1: d.line((x+w+3,216,x+w+20,216),fill=CYAN,width=7)
 im.save(ROOT/'illustrations'/path)
card('systeme_createur.png',['Compétence','Niche','Audience','Contenu','Confiance','Offre utile'])
card('delegation_ia.png',['Humain','Brief','IA / agent','Contrôle humain'])
card('production_video.png',['Recherche','Idée','Brief','Script','Médias','Montage','Publication','Mesure'])
card('mesures.png',['Portée','Intérêt','Demandes','Ventes encaissées'])

# Build structured blocks once for both outputs.
blocks=[]
def add(kind,text='',**kw):blocks.append((kind,text,kw))
add('cover','')
add('title','PAGE DE TITRE')
add('booktitle',TITLE);add('subtitle',SUBTITLE);add('body','Auteur : '+AUTHOR+' · Édition indépendante de lecture · 2026 · ISBN non attribué')
for heading,body in PRELIMS.items():add('newpage');add('h1',heading);add('body',body)
add('newpage');add('h1','TABLE DES MATIÈRES')
for roman,part,lo,hi in PARTS:
 add('tocpart',f'PARTIE {roman} — {part}')
 for c in CHAPTERS[lo-1:hi]: add('tocitem',f"{c['num']:02d}  {c['title']}",anchor=f"chap{c['num']}")
add('tocpart','Conclusion · Annexes · Bibliographie')
add('newpage');add('h1','MODE D’EMPLOI')
add('body','Cette édition de lecture reprend les 39 chapitres et les principales idées du texte communiqué dans la conversation. Elle en propose une adaptation resserrée, et non une transcription intégrale du manuscrit collé. Les sources web du document de départ n’ont pas toutes été revérifiées ; seules les références S01–S13 ci-dessous ont été ouvertes dans cette production. L’auteur doit valider le texte et ses faits biographiques avant diffusion commerciale.')
add('image','systeme_createur.png',caption='De la compétence à une offre utile. Les revenus restent possibles, jamais garantis.')
for roman,part,lo,hi in PARTS:
 add('newpage');add('part',f'PARTIE {roman}');add('parttitle',part)
 if roman=='I':add('image','delegation_ia.png',caption='Un brief et un agent assistent la préparation ; l’humain garde le contrôle.')
 if roman=='IV':add('image','production_video.png',caption='Une chaîne de production documentée, du sujet à la mesure.')
 if roman=='V':add('image','mesures.png',caption='La visibilité, la demande et les ventes encaissées ne sont pas équivalentes.')
 for c in CHAPTERS[lo-1:hi]:
  add('newpage');add('chapter',f"{c['num']:02d}  {c['title']}",anchor=f"chap{c['num']}")
  add('label','OBJECTIF');add('body',c['goal']);add('label','COMPRENDRE');add('body',c['body'])
  add('label','MÉTHODE EN PRATIQUE')
  for i,step in enumerate(c['steps'],1): add('step',f'{i}. {step.strip()}.')
  add('label','EXEMPLE');add('body',c['example'])
  add('box','À retenir — '+c['goal'])
  add('caution',c['caution'])
  add('label','EXERCICE ET CHECKLIST')
  add('body',f"Relisez votre essai : avez-vous vérifié l’entrée, les sources, les droits, le résultat et la prochaine action ? Rédigez votre propre {c['deliverable']} et confrontez-le à une personne concernée.")
  add('label','PROMPT COPIABLE');add('prompt',c['prompt'])
  add('label','PASSAGE À L’ACTION')
  add('body',f"Objectif : {c['goal']} Tâches : suivez la méthode ci-dessus. Checklist : données exactes, sources, droits et validation humaine. Prompt : utilisez le modèle précédent en remplaçant ses variables. Fichier à produire : {c['deliverable']}. Résultat attendu et critère de réussite : {c['success']} Prochaine étape : poursuivez au chapitre suivant après contrôle.")
add('newpage');add('h1','CONCLUSION — CRÉER, VÉRIFIER, APPRENDRE')
add('body','Un système créateur commence par une compétence et un problème observé. L’IA peut accélérer la préparation ; elle ne remplace pas le jugement, la vérification et la responsabilité éditoriale. Les programmes de plateformes sont des canaux, non des garanties. Vérifiez les conditions dans votre pays et à la date où vous agissez. Commencez par une publication pilote, mesurez ce qui correspond à votre objectif, livrez ce que vous promettez et décidez de poursuivre, modifier ou arrêter.')
add('newpage');add('h1','ANNEXE A — BIBLIOTHÈQUE DE PROMPTS')
add('body','Canevas copiables : remplacez les champs entre crochets ; une sortie IA doit être vérifiée. Les données personnelles ne doivent pas être partagées sans nécessité et permission.')
for head,content in ANNEX_PROMPTS:add('h2',head);add('prompt',content)
add('newpage');add('h1','ANNEXE B — MODÈLES DE TRAVAIL')
for title,body in [('Brief','Identifiant · date · responsable · plateforme/pays · public · question · objectif · message · sources · statut des droits · exemple réel ou simulation · durée · CTA · mesure et définition · contrôle humain.'),('Script','SEGMENT_001 · hook honnête · situation · démonstration · limite · CTA · texte à l’écran · source des faits · fichier de voix/vidéo lié.'),('Fiche d’offre','Public · problème · livrable · délais · prix proposé · coûts · exclusions · conditions de livraison et de remboursement · paiement vérifié · suivi.'),('Nommage','SEGMENT_001 → audio_SEGMENT_001 → visual_SEGMENT_001 → sound_SEGMENT_001 → animated_SEGMENT_001 → video_SEGMENT_001. Ajouter _v02 lors d’une correction.')]: add('h2',title);add('body',body)
add('newpage');add('h1','ANNEXE C — PARCOURS EN 30 À 60 JOURS')
for lo,hi,text in CALENDAR:add('h2',f'Jours {lo}–{hi}');add('body',text)
add('body','Le parcours construit des fondations et des processus ; il ne promet pas de ventes. Les jours 31–60 complètent les jours 1–30, ils ne constituent pas un nouveau cycle de soixante jours.')
add('newpage');add('h1','ANNEXE D — CHECKLISTS ET TABLEAU DE SUIVI')
for title,text in [('Avant publication','Question claire ; faits et dates vérifiés ; droits médias ; consentements ; sous-titres ; CTA ; déclaration commerciale/IA pertinente ; export ouvert.'),('Avant l’offre','Besoin observé ; livrable ; délais ; prix ; coûts ; règles locales ; paiement réellement possible.'),('Avant automatisation','Processus manuel stable ; autorisations minimales ; journal ; test réversible ; validation ; bouton d’arrêt.')]:add('h2',title);add('body',text)
add('table','Date | ID | Sujet | Canal | Objectif | Source | Droits | URL | Portée | Demandes | Paiements reçus | Coûts | Décision')
add('newpage');add('h1','ANNEXE E — GLOSSAIRE')
for k,v in GLOSSARY.items():add('h2',k);add('body',v)
add('newpage');add('h1','BIBLIOGRAPHIE ET LIENS OFFICIELS')
add('body','Pages ouvertes le 27 septembre 2026. La présence d’un lien ne confirme pas l’éligibilité d’un pays ou d’un compte. Consultez sources_et_credits.md pour la portée et les réserves. Les 130 références alléguées dans le manuscrit collé n’ont pas toutes été ouvertes pour cette édition.')
for id,title,url,scope in SOURCES: add('source',f'[{id}] {title} — {scope}. Consulté le 27-09-2026. {url}')

# DOCX
D=Document();sec=D.sections[0];sec.page_width=Cm(21);sec.page_height=Cm(29.7);sec.top_margin=Cm(2.25);sec.bottom_margin=Cm(2.1);sec.left_margin=Cm(2.4);sec.right_margin=Cm(2.4)
styles=D.styles;normal=styles['Normal'];normal.font.name='DejaVu Sans';normal.font.size=Pt(10.5);normal.font.color.rgb=RGBColor(24,34,54);normal.paragraph_format.space_after=Pt(8);normal.paragraph_format.line_spacing=1.24
for name,size,color,bold in [('Heading 1',20,NAVY,True),('Heading 2',14,PURPLE,True),('Heading 3',11,CYAN,True)]:
 s=styles[name];s.font.name='DejaVu Sans';s.font.size=Pt(size);s.font.bold=bold;s.font.color.rgb=RGBColor.from_string(color.lstrip('#'));s.paragraph_format.space_before=Pt(12);s.paragraph_format.space_after=Pt(8);s.paragraph_format.keep_with_next=True
for n,sz,col in [('LivreTitre',29,NAVY),('Partie',28,NAVY),('Etiquette',9,CYAN),('Encadre',10,NAVY),('Prompt',9,NAVY),('Legende',9,PURPLE),('Bibliographie',9,NAVY)]:
 s=styles.add_style(n,WD_STYLE_TYPE.PARAGRAPH);s.base_style=normal;s.font.name='DejaVu Sans';s.font.size=Pt(sz);s.font.color.rgb=RGBColor.from_string(col.lstrip('#'));s.font.bold=n in ('LivreTitre','Partie','Etiquette');s.paragraph_format.space_after=Pt(9);s.paragraph_format.keep_with_next=n in ('LivreTitre','Partie','Etiquette')
header=sec.header.paragraphs[0];header.text='BELIVE MONEY  /  SYSTÈME CRÉATEUR IA';header.style=styles['Bibliographie']
footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER;footer.add_run('Valdes Marco BAWA   •   ')
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');footer._p.append(fld)

def bookmark(p,name):
 start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(len(D.paragraphs)+100));start.set(qn('w:name'),name);p._p.insert(0,start)
 end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),start.get(qn('w:id')));p._p.append(end)
def link(p,label,anchor):
 h=OxmlElement('w:hyperlink');h.set(qn('w:anchor'),anchor);run=OxmlElement('w:r');t=OxmlElement('w:t');t.text=label;run.append(t);h.append(run);p._p.append(h)
for kind,text,opts in blocks:
 if kind=='cover':
  p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run().add_picture(str(ROOT/'illustrations/couverture.png'),width=Cm(14.5),height=Cm(20.68))
  p=D.add_paragraph(TITLE,'LivreTitre');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
  p=D.add_paragraph(AUTHOR);p.alignment=WD_ALIGN_PARAGRAPH.CENTER;D.add_page_break();continue
 if kind=='newpage':D.add_page_break();continue
 if kind=='image':
  p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run().add_picture(str(ROOT/'illustrations'/text),width=Cm(16))
  D.add_paragraph(opts.get('caption',''),style='Legende');continue
 if kind=='tocitem':p=D.add_paragraph(style='Normal');link(p,text,opts['anchor']);continue
 mapping={'title':'Heading 1','booktitle':'LivreTitre','subtitle':'Heading 2','part':'Partie','parttitle':'Heading 1','chapter':'Heading 1','h1':'Heading 1','h2':'Heading 2','tocpart':'Heading 2','label':'Etiquette','prompt':'Prompt','box':'Encadre','caution':'Encadre','source':'Bibliographie'}
 p=D.add_paragraph(style=mapping.get(kind,'Normal'))
 if kind=='chapter':bookmark(p,opts['anchor'])
 p.add_run(text)
 if kind in ('prompt','box','caution'):
  shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'EDF5F8' if kind!='caution' else 'FFF5EB');p._p.get_or_add_pPr().append(shade)
 if kind=='source':
  m=re.search(r'https?://\S+',text)
  if m:
   p.runs[0].text=text[:m.start()]
   part=p.add_run(m.group());part.font.color.rgb=RGBColor(0,110,160)
   # External relationship and hyperlink for the bibliography.
   from docx.opc.constants import RELATIONSHIP_TYPE as RT
   rid=p.part.relate_to(m.group(),RT.HYPERLINK,is_external=True)
   h=OxmlElement('w:hyperlink');h.set(qn('r:id'),rid)
   rr=OxmlElement('w:r');tt=OxmlElement('w:t');tt.text=m.group();rr.append(tt);h.append(rr)
   p._p.remove(part._r);p._p.append(h)
 if kind=='table':
  p._element.getparent().remove(p._element)
  t=D.add_table(rows=1,cols=4);t.style='Light Shading Accent 1';t.alignment=WD_TABLE_ALIGNMENT.CENTER
  for cell,v in zip(t.rows[0].cells,['ID / date','Canal / objectif','Preuve / mesure','Décision']):cell.text=v
  for _ in range(3):
   row=t.add_row();[setattr(cell,'text','________________') for cell in row.cells]
# content-type field for native automatic table of contents, in addition to immediately usable linked chapter list
p=D.paragraphs[next(i for i,x in enumerate(D.paragraphs) if x.text=='TABLE DES MATIÈRES')]
toc=D.add_paragraph('Sommaire automatique : dans Word, clic droit → Mettre à jour le champ (après édition).')
p._p.addnext(toc._p)
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'TOC \\o "1-2" \\h \\z \\u');toc._p.append(fld)
D.core_properties.title=TITLE;D.core_properties.author=AUTHOR
D.save(ROOT/'BELIVE_MONEY_manuscrit_final.docx')

# PDF, a *separately typeset* reading edition of same structured blocks.
pdfmetrics.registerFont(TTFont('DejaVu',fontpath));pdfmetrics.registerFont(TTFont('DejaVuBold',boldpath))
S={}
for key,size,lead,col,space in [('body',10,15,NAVY,8),('h1',19,24,NAVY,12),('h2',13,18,PURPLE,8),('chapter',19,25,NAVY,13),('label',9,13,CYAN,6),('prompt',9,14,NAVY,10),('box',10,15,NAVY,7),('caution',10,15,NAVY,7),('source',8,12,NAVY,6),('tocitem',9,14,NAVY,4),('tocpart',12,17,PURPLE,6),('subtitle',13,19,PURPLE,12),('booktitle',26,33,NAVY,16),('part',28,36,NAVY,12),('step',10,15,NAVY,4)]:
 S[key]=ParagraphStyle(key,fontName='DejaVuBold' if key in ('h1','h2','chapter','label','booktitle','part','tocpart') else 'DejaVu',fontSize=size,leading=lead,textColor=colors.HexColor(col),spaceAfter=space,spaceBefore=4 if key in ('h1','h2','label') else 0)
story=[];page_w,page_h=A4
for kind,text,opts in blocks:
 if kind=='newpage':story.append(PageBreak());continue
 if kind=='cover':
  story.append(Image(str(ROOT/'illustrations/couverture.png'),width=350,height=499))
  story.append(Paragraph(html.escape(TITLE),S['booktitle']))
  story.append(Paragraph(AUTHOR,S['h2']));story.append(PageBreak());continue
 if kind=='image':
  story.append(Image(str(ROOT/'illustrations'/text),width=450,height=112.5))
  story.append(Paragraph(html.escape(opts.get('caption','')),S['source']));continue
 if kind=='table':
  rows=[['ID / date','Canal / objectif','Preuve / mesure','Décision']]+[['','','',''] for _ in range(4)]
  t=Table(rows,colWidths=[110,110,120,105],rowHeights=[29]+[32]*4);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor(NAVY)),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,-1),'DejaVu'),('FONTSIZE',(0,0),(-1,-1),8),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#c7d1df'))]));story.append(t);continue
 style=S.get(kind,S['body']);safe=html.escape(text)
 if kind=='tocitem':safe=f'<link href="#{opts["anchor"]}">{safe}</link>'
 if kind=='chapter':safe=f'<a name="{opts["anchor"]}"/>{safe}'
 if kind=='source':safe=re.sub(r'(https?://[^\s&lt;&gt;]+)',lambda m:f'<link href="{m.group(1)}">{m.group(1)}</link>',safe)
 story.append(Paragraph(safe,style))
def decorate(canvas,doc):
 canvas.saveState();canvas.setStrokeColor(colors.HexColor('#d4e3eb'));canvas.line(55,43,540,43)
 canvas.setFont('DejaVu',8);canvas.setFillColor(colors.HexColor(NAVY));canvas.drawString(55,30,'BELIVE MONEY  ·  Édition de lecture');canvas.drawRightString(540,30,str(doc.page));canvas.restoreState()
SimpleDocTemplate(str(ROOT/'BELIVE_MONEY_manuscrit_final.pdf'),pagesize=A4,leftMargin=62,rightMargin=62,topMargin=65,bottomMargin=65,title=TITLE,author=AUTHOR).build(story,onFirstPage=decorate,onLaterPages=decorate)
# maintenance source is an actual structured dump of the content included
with open(ROOT/'manuscrit_source.md','w') as f:
 for kind,text,opts in blocks:
  if kind=='newpage':f.write('\n---\n\n')
  elif kind=='image':f.write(f'\n![{opts.get("caption",text)}](illustrations/{text})\n\n')
  elif kind!='cover':f.write(('# ' if kind in ('h1','chapter','parttitle') else '## ' if kind in ('h2','label') else '')+text+'\n\n')
with open(ROOT/'dashboard_contenu.csv','w',newline='') as f:
 csv.writer(f).writerows([['date','id','sujet','plateforme','objectif','source','droits','lien','portee','demandes','paiements_recus','couts','decision']])
with zipfile.ZipFile(ROOT/'BELIVE_MONEY_manuscrit_final.docx') as z:
 print('DOCX validated:',z.testzip() is None,'embedded media:',[p for p in z.namelist() if p.startswith('word/media/')]);print('TOC field present',b'TOC' in z.read('word/document.xml'))
reader=PdfReader(str(ROOT/'BELIVE_MONEY_manuscrit_final.pdf'))
print('PDF pages',len(reader.pages),'first:',reader.pages[0].extract_text()[:120]);print('last:',reader.pages[-1].extract_text()[-180:]);print('chapters:',len(CHAPTERS))
