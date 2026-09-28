# BELIVE MONEY — QA du diaporama de formation

## Fichier source et comparaison

- Source éditoriale disponible : `BELIVE_MONEY_manuscrit_illustre.docx`, adaptation illustrée et **condensée**, 1 144 paragraphes, 40 tableaux, 8 parties, 39 chapitres, conclusion, annexes A–E, bibliographie S01–S13. Le manuscrit intégral annoncé dans le contexte antérieur **n'est pas disponible** : le support n'en revendique pas la fidélité intégrale.
- PDF compagnon : `BELIVE_MONEY_manuscrit_illustre.pdf`, 69 pages ; rendu de lecture **composé séparément**, pas un export Word. Son extraction de texte présente des substitutions de glyphes accentués et des avertissements de police/flux compressé sous PyMuPDF/pypdf. Les 39 têtes de chapitre sont retrouvées dans l'extraction PDF sous ces substitutions. Le PDF sert donc de repère d'ordre, **pas** de texte de référence ; le DOCX prime. Aucune comparaison binaire mot à mot fiable n'est revendiquée. Aucun `Slide2.pdf` n'a été fourni ni consulté.
- Lecture structurée : pour chacun des 39 chapitres, le constructeur `slides/content.py` extrait titre, objectif, explication, méthode complète, exemple et prompt copiable. Il échoue si un chapitre ou une de ces rubriques manque. Les exemples restent attribués au livre et désignés comme simulations en notes ; les prompts de six chapitres sélectionnés sont reportés intégralement dans les notes des slides de prompt. Annexes B/C/D prises en compte pour modèles, calendrier et checklist ; conclusion pour la clôture ; bibliographie pour les sources. Le deck reste une **sélection pédagogique** et non la reproduction intégrale du DOCX.

## Artefacts créés et vérifiés

| Élément | Résultat |
|---|---|
| `BELIVE_MONEY_formation_belive_money.pptx` | PPTX natif éditable au format 16:9 (13,333 × 7,5 pouces), 84 diapositives, 179 min prévues ; 1 726 formes dont cartes, textes, flèches et cases. 84 slides ont des notes intégrées. |
| `BELIVE_MONEY_storyboard_presentation.md` | 84 fiches de slide : identifiant, partie, objectif, idée, texte, exemple/exercice, notes, visuel, chapitre, source et durée. |
| `BELIVE_MONEY_notes_presentateur.md` | Reprise des notes de chaque slide, prompts complets inclus. |
| `BELIVE_MONEY_plan_illustrations_slides.md` | Placement, justification des formes natives et alt text pour 84 slides. |
| `BELIVE_MONEY_credits_sources_images.md` | Sources officielles, limites, date de consultation, crédits d'illustration ; aucune photo ou image externe. |
| `BELIVE_MONEY_apercus_slides/` | 84 PNG d'aperçu de contrôle **interne**, reconstruits à partir des éléments du PPTX. Ce n'est pas un rendu PowerPoint. Aucun PDF du deck n'est livré, faute de moteur d'export. |

## Vérifications exécutées

1. Construction : `.venv/bin/python production/build_deck.py` → 84 diapositives ; `.venv/bin/python production/build_slide_reports.py` → 84 fiches et 179 min ; `.venv/bin/python production/verify_deck.py` → **CONFORME**.
2. PPTX rouvert avec `python-pptx` : 84 diapositives ; ratio 1,7777 ; 0 note manquante ; 0 forme hors cadre ; 0 texte alternatif absent ; 0 débordement **selon une estimation métrique** ; police déclarée Arial ; 0 lien externe cliquable (URLs fournies dans le registre). Les notes totalisent 48 289 caractères ; les codes S001–S084 ont été retrouvés dans les notes.
3. Ordre contrôlé : huit parties et 39 chapitres dans leur ordre, puis trois blocs transversaux explicitement marqués (prompts, vérification, mise en œuvre). Chaque chapitre a sa méthode visuelle en éléments éditables. Processus à huit décisions sur deux lignes numérotées ; workflow ch. 17 en cinq étapes. Schémas originaux sans statistiques ni logo.
4. Aperçus : `.venv/bin/python production/render_deck_preview.py` → 84 PNG de contrôle interne. Aperçus S003 (fil conducteur), S005 (premier chapitre), S031 (workflow), S067 (intertitre pratique), S077 (YouTube) ouverts dans le visualiseur ; planche-contact des 84 aperçus parcourue ; les cartes et titres sont lisibles, et le fil conducteur a été corrigé en deux rangées pour éviter des flèches sur le texte. Des contrôles automatiques couvrent les 84 géométries ; **aucune inspection visuelle humaine exhaustive diapo par diapo ni sortie physique projecteur** n'a été réalisée.
5. Conditions évolutives : pages officielles TikTok T1/T2, YouTube Y1/Y2, Meta M1 et Gumroad G1 ouvertes le 28-09-2026 ; dates et limites dans le registre. Le deck évite toute promesse d'admissibilité ou de revenu. Les critères futurs YouTube (2027) sont identifiés comme annonce future.

## Limites et actions avant présentation publique

- Aucun logiciel PowerPoint ni LibreOffice n'est disponible ici : **ouverture et rendu sous PowerPoint, animations, lecture des notes, comportement des liens, substitution de polices, contraste sur projecteur et export PDF PowerPoint restent à vérifier chez le destinataire.** La vérification d'un débordement est une estimation, pas une validation visuelle PowerPoint.
- Aucun PDF de présentation n'a été exporté ni livré : aucun moteur PowerPoint/LibreOffice n'est disponible. Les PNG de contrôle sont des reconstructions indépendantes, non des épreuves d'impression.
- Le PDF du livre présente des anomalies d'extraction (glyphes/avertissements de décompression) ; consulter le DOCX pour relire les citations et légendes. La concordance exacte de chaque mot entre le PDF et le DOCX n'a pas pu être validée automatiquement.
- Les 84 diapositives et 179 minutes constituent une formation longue : prévoir deux ou trois séances ou sélectionner des modules. Aucune limite arbitraire de diapositives n'a été imposée ; les contenus détaillés sont en notes.
- Les pages de conditions d'outils/plateformes et de paiement peuvent changer dès après le 28-09-2026. Recontrôler pays, programme, compte et possibilité réelle de recevoir/retirer avant tout usage pratique. Les juridictions et règles fiscales sont à valider localement.
- Faire une relecture éditoriale finale et un test avec PowerPoint sur l'appareil de diffusion ; le texte alternatif OOXML est renseigné, mais son exposition effective par tous lecteurs d'écran n'a pas été testée.
