# BELIVE MONEY — édition de lecture illustrée

**Source réellement accessible :** `BELIVE_MONEY_manuscrit_final.docx`, une édition antérieure **adaptée et condensée** des 39 chapitres. Le document `BELIVE_MONEY_manuscrit_integral.docx` annoncé (~114 pages) n'est pas présent. Ne pas présenter cette version comme une illustration de l'intégral.

## Nouveaux livrables

- [BELIVE_MONEY_manuscrit_illustre.docx](BELIVE_MONEY_manuscrit_illustre.docx) — 39 tableaux pédagogiques Word natifs ajoutés et huit nouvelles ouvertures de partie ; les paragraphes source sont conservés en séquence.
- [BELIVE_MONEY_manuscrit_illustre.pdf](BELIVE_MONEY_manuscrit_illustre.pdf) — PDF de lecture composé **séparément** à partir des blocs du DOCX (pas un export Word).
- [BELIVE_MONEY_plan_visuels.md](BELIVE_MONEY_plan_visuels.md), [BELIVE_MONEY_prompts_images_et_sources.md](BELIVE_MONEY_prompts_images_et_sources.md), [BELIVE_MONEY_QA_visuels.md](BELIVE_MONEY_QA_visuels.md), [BELIVE_MONEY_journal_modifications.md](BELIVE_MONEY_journal_modifications.md).
- [BELIVE_MONEY_illustrations/](BELIVE_MONEY_illustrations/) — neuf fichiers PNG utilisés, dont la couverture existante.

Les contrôles et limites avant diffusion se trouvent dans le QA et le journal. L'ancienne [édition adaptée](BELIVE_MONEY_manuscrit_final.pdf) et son ZIP restent dans ce dépôt comme base, non comme manuscrit intégral.

## Reconstruction

Créer un environnement Python avec `python-docx`, `Pillow`, `reportlab`, `pypdf`. Lancer `python production/illustrate.py`, puis `python production/render_illustrated_pdf.py` et `python production/build_visual_reports.py`. Les schémas sont décrits dans `production/visual_specs.py`.
