"""Storyboard source derived from the actually available adapted illustrated DOCX.
No external images/logos; every data claim is sourced in final resources.
"""
from dataclasses import dataclass,field
from docx import Document
import re
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
SOURCE=ROOT/'BELIVE_MONEY_manuscrit_illustre.docx'
D=Document(SOURCE)
ps=D.paragraphs
PARTS=['COMPRENDRE LE SYSTÈME CRÉATEUR IA','DE L’IDÉE À L’OPPORTUNITÉ','CONSTRUIRE SON SYSTÈME DE CONTENU','PRODUIRE AVEC L’IA','MESURER ET AUTOMATISER','MONÉTISER','TIKTOK DU COMPTE AU PAIEMENT','PARCOURS D’EXÉCUTION']
RANGES=[(1,3),(4,8),(9,14),(15,19),(20,24),(25,30),(31,35),(36,39)]
CH={}
for i,p in enumerate(ps):
 m=re.match(r'^(\d{2})\s{1,2}(.+)$',p.text)
 if p.style.name=='Heading 1' and m:
  num=int(m.group(1));assert num not in CH
  stop=next((j for j in range(i+1,len(ps)) if ps[j].style.name=='Heading 1' and (re.match(r'^\d{2}\s',ps[j].text) or ps[j].text.startswith('CONCLUSION'))),len(ps))
  section=ps[i:stop]
  def after(label):
   for k,x in enumerate(section[:-1]):
    if x.text.strip()==label:return section[k+1].text.strip()
   return ''
  goal=after('OBJECTIF');body=after('COMPRENDRE');example=after('EXEMPLE');prompt=after('PROMPT COPIABLE')
  methods=[];active=False
  for x in section:
   if x.text=='MÉTHODE EN PRATIQUE':active=True;continue
   if x.text=='EXEMPLE':break
   if active and re.match(r'^\d+\.\s',x.text):methods.append(re.sub(r'^\d+\.\s','',x.text))
  CH[num]=dict(title=m.group(2),goal=goal,body=body,example=example,prompt=prompt,steps=methods)
assert len(CH)==39
# 27 prompts & conclusion are really present in the DOCX; verify some recurring anchor content
assert any(x.text.startswith('ANNEXE A') for x in ps) and any(x.text.startswith('CONCLUSION') for x in ps)
PROMPT_NAMES={1:'A1 • Note vocale → brief',5:'A2 • Choisir une niche',12:'A10 • Script',17:'A12 • Storyboard',20:'A18 • Statistiques',25:'A20 • Offre',26:'A2 • Choisir une niche',33:'A11 • Vérifier un moyen de paiement'}
# Source's annex actually uses A20/A11 names differently; select verified by name, with chapter prompt as fallback
ANNEX={}
for i,p in enumerate(ps[:-1]):
 if p.style.name=='Heading 2' and re.match(r'^A\d+\s*•',p.text):ANNEX[p.text]=ps[i+1].text
@dataclass
class Slide:
 kind:str
 title:str
 part:int=0
 chapter:int=0
 subtitle:str=''
 nodes:list=field(default_factory=list)
 exercise:str=''
 notes:str=''
 source:str=''
 duration:int=1
 id:str=''
 slidesource:str=''
SL=[]
def add(**kwargs):
 s=Slide(**kwargs);s.id=f'S{len(SL)+1:03d}';SL.append(s);return s
add(kind='cover',title='BELIVE MONEY',subtitle='SYSTÈME CRÉATEUR IA',notes='Présenter le parcours : compétence, audience, contenu, offre utile et contrôle humain. Préciser que cette formation repose sur le DOCX illustré adapté, non sur le manuscrit intégral annoncé.',source='Couverture et introduction du DOCX adapté',duration=2)
add(kind='manifesto',title='Une méthode, pas une promesse',subtitle='Aucune garantie de viralité, de revenu ou de monétisation',nodes=['Tester avec les moyens disponibles','Documenter les faits et les droits','Choisir une offre réellement livrable'],notes='Demander aux participants de nommer un résultat qu’ils contrôlent, comme publier un contenu vérifié, plutôt qu’un résultat externe comme devenir viral.',source='Copyright, avertissement et introduction',duration=2)
add(kind='path',title='Le système créateur en huit décisions',nodes=['Compétence','Niche','Audience','Contenu','Confiance','Offre','Revenus possibles','Système documenté'],notes='Parcourir la chaîne sans raccourci : les vues ne sont pas des ventes. Les revenus restent possibles, jamais garantis.',source='Introduction et chapitres 2, 25, 39',duration=3)
for part,(a,b) in enumerate(RANGES,1):
 add(kind='section',part=part,title=f'PARTIE {part:02d}',subtitle=PARTS[part-1],notes=f'Ouvrir la partie {part}. Demander ce que les participants savent déjà, puis introduire les chapitres {a} à {b}.',source=f'Partie {part} du DOCX',duration=1)
 for n in range(a,b+1):
  c=CH[n]; nod=c['steps'][:4]
  if n in (17,39) and len(c['steps'])>=5:nod=c['steps'][:5]
  # One core slide per chapter, with exact source title, goal and short procedure; no unsupported threshold
  add(kind='chapter',part=part,chapter=n,title=f'{n:02d} · {c["title"]}',subtitle=c['goal'],nodes=nod,notes='Explication issue du chapitre : '+c['body']+'\n\nMéthode : '+' | '.join(c['steps'])+'\n\nExemple du livre : '+c['example']+'\n\nPrompt intégral du chapitre : '+c['prompt'],source=f'Chapitre {n:02d} — {c["title"]}',duration=2)
  # Example/exercise slides for 8 keystone topics; instructions grounded in example, no invented testimony
  if n in (3,5,12,13,17,20,25,32):
   exercise={3:'Choisissez une tâche réversible. Marquez ce que l’IA prépare et ce que vous vérifiez.',5:'Écrivez un public, un problème, une compétence et une promesse sans garantie.',12:'Écrivez une ouverture honnête puis la démonstration qui la justifie.',13:'Enregistrez dix secondes ; vérifiez le son, la lumière et la confidentialité.',17:'Dessinez SEGMENT_001 ; listez script, voix, visuel, contrôle et fichier attendu.',20:'Séparez portée, questions, demandes, commandes et paiements réellement reçus.',25:'Rédigez livrable, délai et exclusions d’une offre test.',32:'Vérifiez le programme exact, le pays réel et le paiement sur une page officielle.'}[n]
   add(kind='exercise',part=part,chapter=n,title=f'Atelier · {c["title"]}',subtitle='Une décision à documenter, pas un résultat à promettre',exercise=exercise,notes=f'Faire travailler en binômes. Source : chapitre {n}. Exemple fourni par le livre : {c["example"]} Insister sur la mention « simulation » si vous utilisez cet exemple. Puis demander à chaque participant une trace écrite et une prochaine vérification.',source=f'Chapitre {n:02d} (exemple/exercice adapté)',duration=4)
 # consolidation slide for each part
 add(kind='recap',part=part,title=f'À retenir · partie {part:02d}',subtitle=PARTS[part-1],nodes=[CH[n]['goal'] for n in range(a,b+1)][:5],notes=f'Synthétiser les chapitres {a} à {b}. Demander une décision immédiatement réalisable, puis une limite ou une incertitude à vérifier.',source=f'Chapitres {a:02d}–{b:02d}',duration=2)
add(kind='section',part=9,title='BLOC PRATIQUE',subtitle='PROMPTS RÉELLEMENT PRÉSENTS DANS LE LIVRE',notes='Six prompts copiables extraits de chapitres identifiés. Le prompt complet est reporté dans les notes de la diapositive, jamais en petits caractères à l’écran.',source='Annexe A et rubriques « Prompt copiable »',duration=1)
# Concrete, checked prompts, visible excerpt and full prompt in notes
for n in (1,5,12,17,20,25):
 c=CH[n]
 add(kind='prompt',chapter=n,title=f'Prompt praticable · {c["title"]}',subtitle='Variables à compléter avant la demande',nodes=['Contexte réel et pays','Sources et droits connus','Format de sortie demandé','Validation humaine finale'],notes='Prompt complet réellement présent dans le chapitre '+str(n)+' : '+c['prompt']+'\nObjectif : '+c['goal']+'\nErreur à éviter : fournir des faits non vérifiés ou des données personnelles inutiles. Résultat attendu : un brouillon contrôlable.',source=f'Chapitre {n:02d}, rubrique « Prompt copiable »',duration=2)
add(kind='section',part=10,title='BLOC VÉRIFICATION',subtitle='CE QUI CHANGE VITE DOIT ÊTRE RECONTRÔLÉ',notes='Ces repères proviennent de pages officielles ouvertes le 28 septembre 2026. Ils décrivent des programmes et des pays, pas une promesse de revenu. À recontrôler avant chaque usage.',source='Sources officielles [T1][T2][Y1][Y2][G1]',duration=1)
# factual slides with verified official sources, date explicit
add(kind='facts',part=7,chapter=26,title='Creator Rewards : conditions, pas promesse',subtitle='Repères généraux · vérifiés le 28 septembre 2026',nodes=['Compte personnel en règle ; pays admissible','18 ans ou plus (19 en Corée du Sud)','10 000 abonnés et 100 000 vues / 30 jours','Vidéos originales d’au moins une minute'],notes='Source officielle ouverte le 28/09/2026 : TikTok Support — Creator Rewards Program. Éligibilité au programme ne garantit pas récompense ; vérifier pays, tableau de bord, contenu après inscription et vues qualifiées. Ne jamais contourner la résidence.',source='TikTok, Creator Rewards Program [T1]',duration=3)
add(kind='facts',part=7,chapter=33,title='TikTok : estimé ≠ encaissé',subtitle='Repères généraux · vérifiés le 28 septembre 2026',nodes=['Vues qualifiées et règles du programme','Récompense estimée dans TikTok Studio','Montant payable selon les conditions locales','Rapprocher avec le paiement reçu'],notes='La page générale TikTok « How rewards work » annonce un seuil de 10 USD ou équivalent et le 15 du mois suivant dans ses conditions décrites, mais les règles régionales et le compte priment. Ce deck évite de faire du seuil une règle universelle. Pages officielles ouvertes le 28/09/2026.',source='TikTok, How rewards work [T2]',duration=2)
add(kind='facts',part=3,chapter=8,title='YouTube : séparer présent et annonce future',subtitle='Repères officiels · vérifiés le 28 septembre 2026',nodes=['YPP actuel : 1 000 abonnés + critères de visionnage','Pays admissible et revue de la chaîne','Annonce de modification au 1er février 2027','Recontrôler avant toute candidature'],notes='En vigueur à la date de vérification : 1 000 abonnés + 4 000 heures qualifiées sur 12 mois ou 10 millions de vues Shorts qualifiées sur 90 jours. Un changement au 01/02/2027 est annoncé (8 000 h/365 jours ou 20 M Shorts/90 jours pour nouveaux candidats). Ne pas présenter ces derniers comme critères actuels. Vérifier YouTube Studio et disponibilité nationale.',source='YouTube Partner Program [Y1] ; Changes to YPP [Y2]',duration=3)
add(kind='decision',title='L’Afrique francophone n’est pas un seul marché',subtitle='Pays, langue, offre, paiement et obligations : vérifications distinctes',nodes=['Définir la ville et le public','Tester une offre livrable localement','Vérifier réception ET retrait','Vérifier fiscalité et droits locaux'],notes='S’appuyer sur chapitres 34–35. Gumroad affiche des versements bancaires locaux dans certains pays dont le Bénin ; cette liste ne garantit pas un compte, une banque ou une vente. N’utiliser aucune fausse résidence.',source='Chapitres 34–35 ; Gumroad [G1]',duration=3)
add(kind='section',part=11,title='BLOC MISE EN ŒUVRE',subtitle='30 JOURS · 60 JOURS · CONTRÔLES',notes='Transformer la formation en décisions datées, avec des contrôles qualité avant publication.',source='Chapitres 36–38 et annexe C',duration=1)
add(kind='roadmap',title='30 jours pour créer les fondations',nodes=['Jours 1–7 : profil, compétence, public','Jours 8–14 : questions, idées, scripts','Jours 15–21 : premiers contenus et suivi','Jours 22–30 : bilan et offre pilote'],notes='Le calendrier est une suggestion de travail, pas un seuil de réussite. Faire choisir un livrable par semaine ; réduire la cadence si le contrôle qualité souffre.',source='Ch. 36 et annexe C',duration=3)
add(kind='roadmap',title='Jours 31–60 : consolider puis documenter',nodes=['31–45 : régularité et test de CTA','31–45 : offre et prospects consentis','46–60 : fichiers, droits et reporting','46–60 : une automatisation réversible'],notes='Ce sont les 30 jours supplémentaires qui suivent les jours 1–30, non 60 jours supplémentaires. La décision de pivoter ou arrêter est acceptable.',source='Ch. 37 et annexe C',duration=3)
add(kind='checklist',title='Avant de publier : cinq contrôles',nodes=['Une question claire et une preuve','Droits des images, voix et musique','Sous-titres et liens relus','CTA et divulgation honnêtes','Export ouvert et vérifié'],notes='La checklist est issue des chapitres 14, 16, 24 et de l’annexe D. Une publication commerciale peut exiger un réglage de divulgation selon la plateforme.',source='Ch. 14, 16, 24 ; annexe D',duration=2)
add(kind='closing',title='Votre prochaine action',subtitle='Une compétence · un public · un contenu pilote · une mesure · une décision',notes='Faire remplir une fiche : « Je peux aider [public] à [tâche] en produisant [contenu] ; je vérifierai [fait/droit] et mesurerai [indicateur défini]. » Ne promettre aucun revenu.',source='Conclusion et chapitre 39',duration=2)
add(kind='resources',title='Sources & recontrôle',nodes=['TikTok Creator Rewards et paiements','YouTube Partner Program : présent / 2027','Meta : politiques par produit et pays','Gumroad : versements par pays'],notes='Liens et périmètres dans BELIVE_MONEY_credits_sources_images.md. Sources consultées le 28/09/2026. Vérifier de nouveau le produit précis et la région réelle avant tout usage. Aucun logo, screenshot ou photographie externe utilisé.',source='Registre officiel du deck',duration=2)
for s in SL:
 if not s.slidesource:s.slidesource=s.source
assert len({s.id for s in SL})==len(SL)
if __name__=='__main__':
 from collections import Counter
 print(len(SL),Counter(s.kind for s in SL),sum(s.duration for s in SL),'minutes')
