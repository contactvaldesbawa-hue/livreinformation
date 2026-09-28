from pathlib import Path
from visual_specs import CONCEPTS,PART_THEMES
from docx import Document
from PIL import Image
import csv
ROOT=Path(__file__).resolve().parent.parent
src=Document(ROOT/'BELIVE_MONEY_manuscrit_final.docx')
chapters=[p.text for p in src.paragraphs if p.style.name=='Heading 1' and p.text[:2].strip().isdigit()]
PARTS=[(1,3),(4,8),(9,14),(15,19),(20,24),(25,30),(31,35),(36,39)]
def part(i):return next(j for j,(a,b) in enumerate(PARTS,1) if a<=i<=b)
plan=['# BELIVE MONEY — plan des visuels de l’édition illustrée','', 'Source inspectée : `BELIVE_MONEY_manuscrit_final.docx` (édition adaptée, non le manuscrit intégral annoncé). Document de départ : 1097 paragraphes, huit parties, 39 chapitres, cinq images incorporées et un tableau ; aucune pagination fiable avant rendu. Cette édition nouvelle préserve l’ordre des paragraphes de ce DOCX.','', 'Une palette bleu nuit, cyan et violet a été proposée pour ce projet ; elle ne constitue pas une charte officielle. Les diagrammes de chapitre sont des tableaux Word natifs éditables, non des SmartArt. Les cartes ne comportent aucun logo, chiffre prétendument réel ni interface. Leurs labels constituent une synthèse pédagogique de chaque chapitre.','', '| ID | Partie / section | Objectif | Type / composition | Ratio / résolution | Source ou méthode / texte visible | Crédit et droits | Alt text | Emplacement Word | Statut et contrôle |','|---|---|---|---|---|---|---|---|---|---|']
qa=['# BELIVE MONEY — QA visuels','', 'Édition illustrée **adaptée**, pas le manuscrit intégral en 114 pages annoncé. Contrôles techniques effectués le 2026-09-28 ; les rendus Word ne peuvent pas être ouverts avec Word/LibreOffice dans cet environnement. Le PDF est composé séparément depuis les éléments du DOCX, et non exporté par Word.','', '| ID | Statut | Emplacement et droit | Accessibilité / intégration | Contrôle effectué / réserve |','|---|---|---|---|---|']
for i,(title,goal,labels) in enumerate(CONCEPTS,1):
 id=f'C{i:02d}';location=f'Partie {part(i)} / chapitre {i:02d}'
 plan.append(f'| {id} | {location} : {chapters[i-1]} ; après le titre | {goal} | Tableau Word natif à deux lignes ; étapes : {" → ".join(labels)} | largeur cible 16 cm / sans raster | Composition déterministe, libellés : {" ; ".join(labels)} | Création originale, aucune source externe ; aucun logo | {goal}. Étapes : {", ".join(labels)} | Après le Heading 1 du chapitre {i:02d}, avant OBJECTIF | Créé ; tableau relu et présence vérifiée dans DOCX/PDF |')
 qa.append(f'| {id} | Intégré | Après le titre du chapitre {i:02d} ; tableau original, sans asset tiers | Libellés natifs lisibles ; légende et description textuelle. Alt text non applicable aux cellules ; ordre de lecture Word à tester | Présent dans DOCX et PDF ; contrôle visuel ponctuel, pas 39 pages inspectées individuellement |')
for i,label in enumerate(PART_THEMES,1):
 id=f'P{i:02d}';fn=f'partie_{i:02d}.png'
 plan.append(f'| {id} | Ouverture partie {i} | {label} | Cinq cartes reliées ; disposition unique, sans texte dans image | 1800×430 px, bandeau ~4,2:1 | Dessin Pillow déterministe, sans visuel tiers ; légende « Ouverture de partie {i} — {label} » | Création originale interne | Cinq étapes reliées de la partie {i}, sans logo ni interface | Après le titre de la partie {i}, avant le premier chapitre | Créé, incorporé au DOCX et visible dans le PDF (sondage) |')
 qa.append(f'| {id} | Intégré | Partie {i} ; `BELIVE_MONEY_illustrations/{fn}`, origine interne | Alt text OOXML défini sur l’image nouvelle ; légende Word | Dimensions 1800×430 ; image incorporée ; vérifier visuellement chaque page dans Word |')
plan.append('| COVER | Couverture | Métaphore du système créateur solo | Main avec téléphone et nœuds lumineux ; aucun texte dans l’image | 864×1232 px, portrait | Prompt exact dans le journal des sources, titre Word séparé | Image IA Arena ; droits de diffusion commerciale à confirmer | Main tenant un téléphone, chemin lumineux vers des étapes reliées | Couverture existante conservée | Préexistante dans DOCX source ; copie autonome ; faible résolution pour prépresse |')
qa.append('| COVER | Conservée | Couverture préexistante, `BELIVE_MONEY_illustrations/couverture.png` | Description OOXML générique ajoutée aux images antérieures ; texte du titre reste natif | Vue en PDF page 1 ; 864×1232 insuffisant pour couverture pleine page à 300 ppp |')
plan.extend(['','## Ressources déjà incorporées dans le DOCX source','', 'La source contenait cinq médias PNG (couverture et quatre schémas génériques), que la nouvelle édition conserve sans les déplacer. Ces images ne doivent pas être comptées comme les huit nouvelles images d’ouverture de parties. Quatre schémas nouveaux utilisables sont des **tableaux Word**, pas des fichiers image autonomes : la consigne « un fichier par image » ne s’applique qu’aux neuf fichiers PNG utilisés listés ci-dessus. Le tableau de suivi Word préexistant est conservé.','', '## Choix de méthode','', 'Aucune photo trouvée en ligne n’est utilisée : la clarification des étapes bénéficie davantage de matrices lisibles et éditables, sans droits des personnes ou des marques. Aucun graphique chiffré ni capture d’interface. La page de couverture n’a ni dos ni fond perdu ; édition de lecture seulement.'])
qa.extend(['','## Compte et limitations','','- 39/39 chapitres illustrés par de nouveaux tableaux Word natifs (une matrice par chapitre).','- 8/8 parties possèdent un bandeau PNG inédit ; couverture IA préexistante conservée.','- 0 photographie externe ; 1 image générée par IA (préexistante) ; 8 nouveaux bandeaux conceptuels dessinés ; 39 nouveaux tableaux-diagrammes éditables ; 0 graphique quantitatif ; 1 ancien tableau de suivi et 4 anciens schémas conservés.','- Images réellement intégrées dans le DOCX final : 11 médias OOXML uniques (les cinq préexistants, plus huit nouveaux bandeaux distincts).','- Le texte des 1097 paragraphes originaux est conservé en séquence dans le DOCX. Cela ne signifie PAS que le long manuscrit envoyé en conversation est reproduit dans le document de départ.','- PDF : rendu séparé issu du DOCX ; ne pas qualifier d’export Word. Les 69 pages ont été techniquement rasterisées ; pages 20, 46 et 52 inspectées visuellement ; relecture éditoriale complète reste nécessaire.','- Images du DOCX source avec description alternative générique ; nouveau bandeau avec description OOXML spécifique. Titres, labels et légendes sont sélectionnables.'])
(ROOT/'BELIVE_MONEY_plan_visuels.md').write_text('\n'.join(plan)+'\n')
(ROOT/'BELIVE_MONEY_QA_visuels.md').write_text('\n'.join(qa)+'\n')
source='''# BELIVE MONEY — prompts d’images et journal des droits

Contrôle le 2026-09-28. Aucune photographie externe ni logo n’a été téléchargé. Les fichiers `BELIVE_MONEY_illustrations/partie_01.png` à `partie_08.png` sont dessinés par `production/illustrate.py` (Pillow), sans appel à un générateur d’images. Leur composition exacte : fond bleu nuit #101A35, cinq cartes bleu/violet reliées par traits cyan, disposition trigonométrique variable par partie, pictogrammes géométriques abstraits ; 1800×430 px ; pas de texte dans l’image. Les titres des parties restent du texte Word. Licence d’éléments tiers : sans objet pour ces huit créations originales. Crédit proposé : « Illustration conceptuelle originale BELIVE MONEY ».

## Couverture existante, conservée

Outil connu : Arena, génération d’image lors d’une session précédente. Prompt exact consigné lors de cette création :

> Premium editorial book cover illustration only, no text, no letters, no logos. Portrait A4 ratio. Deep midnight navy background, elegant translucent cyan and subtle violet luminous pathways flowing from a human hand holding a simple smartphone at lower left into a constellation of connected circular nodes and upward structured steps, sophisticated geometric vector-meets-3D editorial style, clean strong negative space at top half reserved for title overlay, modern francophone business nonfiction aesthetic, no money symbols, no platform interface.

Nom final : `BELIVE_MONEY_illustrations/couverture.png` (copie identique à `illustrations/couverture.png`). Aucun texte incorporé dans l’image ; le titre est du texte natif. Aucun logo ni visage réel. Les conditions d’exploitation commerciale de cette image IA et sa qualité d’impression à 300 ppp demandent validation humaine. La présente note ne constitue pas une licence autonome.

## Schémas éditables C01–C39

La méthode retenue n’est **pas** la génération d’image. Les libellés français et l’ordre des étapes figurent dans `production/visual_specs.py` et dans le plan. Construction : tableau OOXML natif de deux lignes, fond bleu nuit pour l’index, alternance bleu clair/violet doux pour les labels ; largeur prévue de 16 cm ; une légende native par chapitre. Source des concepts : titres et idées de `BELIVE_MONEY_manuscrit_final.docx`, adaptation précédemment créée, sans chiffre ou fait externe. Pas de licence d’image à invoquer pour ces cellules. **Pas de fichier image séparé pour un tableau Word éditable.**

## Réemploi des schémas déjà présents dans le DOCX source

Les quatre schémas PNG de l’édition adaptée, générés par code Pillow lors de la phase précédente, restent dans le document source sans modification de leurs pixels. Crédits et limitations : `sources_et_credits.md` et `plan_illustrations.md`. Aucune nouvelle URL externe ni droit sur un actif tiers présumé.

## Plateformes et marques

Les pages officielles mentionnées dans le livre sont des **sources factuelles**, non des licences de visuel. Aucun logo, charte de marque, capture ou interface d’une plateforme n’a été utilisé. Aucun partenariat n’est suggéré.
'''
(ROOT/'BELIVE_MONEY_prompts_images_et_sources.md').write_text(source)
