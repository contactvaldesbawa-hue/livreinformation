# Journal des modifications — édition illustrée, 28 septembre 2026

## Source réelle et limites

- Aucun `BELIVE_MONEY_manuscrit_integral.docx` de 114 pages n’est accessible dans le dépôt ou les fichiers joints à cette session. Source de production : **`BELIVE_MONEY_manuscrit_final.docx`**, édition antérieure **adaptée et condensée**, 1097 paragraphes, 1 tableau, 5 images incorporées. Le texte intégral collé précédemment en conversation n’est **pas** recréé ici ; aucun titre de fichier ne doit faire croire le contraire.
- Audit de la source : huit grandes parties et 39 chapitres, pages liminaires, conclusion, annexes et bibliographie. La pagination d’origine du DOCX n’est pas vérifiable sans moteur Word/LibreOffice.

## Travail réalisé

1. Sans écraser le DOCX source, création de `BELIVE_MONEY_manuscrit_illustre.docx` : les **1097 paragraphes source sont conservés en ordre**, les cinq images et le tableau source restent présents.
2. Après chaque titre de chapitre, insertion d’une **matrice de procédure Word native éditable** (39 nouvelles tables) et d’une légende. Les étapes sont synthétisées du contenu correspondant ; aucun chiffre réel inventé.
3. Ajout de huit bandeaux PNG originaux, distincts, aux ouvertures de partie. Aucun logo, capture d’écran, personne identifiable ou interface factice. Conservation de la couverture IA antérieure sans retoucher son contenu ; nom d’auteur et titre restent en texte Word.
4. Ajout d’un texte alternatif OOXML spécifique aux nouveaux bandeaux, d’une description générique pour les cinq images source non décrites et de légendes textuelles. Les tableaux Word natifs sont lus comme texte ; l’ordre de lecture effectif doit être testé dans Word et avec une technologie d’assistance.
5. Composition indépendante de `BELIVE_MONEY_manuscrit_illustre.pdf` **à partir du DOCX illustré** par un moteur Python/ReportLab : ce fichier de 69 pages n’est **pas un export du moteur Word**. Tous les paragraphes, tableaux et images ont été parcours dans l’ordre ; la pagination et la reproduction des liens/notes peuvent diverger de Word.
6. QA : comparaison algorithmique de la séquence des paragraphes source avec le DOCX final ; 39 titres, 40 tables au total, 13 médias OOXML (5 anciens + 8 nouveaux), 69 pages PDF. Chaque page PDF a été rasterisée techniquement ; échantillon des pages 20, 46 et 52 ouvert visuellement. Aucun graphe chiffré ni photo externe ajouté.

## Points non résolus avant publication

- Le souhait d’illustrer le **manuscrit intégral de 114 pages** requiert le vrai fichier intégral ; la source de cette livraison est condensée, et la fidélité au texte intégral ne peut être attestée.
- Vérifier le document final dans Word/LibreOffice, notamment les sauts, coupures, tailles de tableau, en-têtes, pagination, liens et champ de sommaire automatique. Le système ne contient ni Word ni LibreOffice, leur installation via apt a échoué (dépôt Debian inaccessible).
- Les schémas de chapitres sont une famille de matrices sobres cohérentes plutôt que 39 photographies/illustrations uniques. Si une direction artistique plus narrative est souhaitée, concevoir une nouvelle passe après réception de l’intégral.
- Couverture 864×1232 px, non qualifiée pour prépresse A4 pleine page à 300 ppp ; aucun fond perdu, dos ni gabarit d’imprimeur. Vérifier l’exploitation commerciale de l’image IA et les faits évolutifs du livre avant vente.
