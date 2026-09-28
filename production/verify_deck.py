#!/usr/bin/env python3
"""Contrôles automatiques du PPTX livré.

Vérifie ce qui est réellement vérifiable sans PowerPoint/LibreOffice :
nombre et ordre des diapositives, texte alt sur chaque forme, notes de
présentateur, débordements estimés du texte, formes hors zone, polices,
couleurs et liens. Le rendu visuel est traité par render_deck_preview.py.
"""
from pathlib import Path
import json
import sys

from pptx import Presentation
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from slides.metrics import n_lines  # noqa: E402

PPTX = ROOT / "BELIVE_MONEY_formation_belive_money.pptx"
EMU_IN = 914400.0
prs = Presentation(PPTX)
SW, SH = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
TXT_TYPES = {17, 19}  # TEXT_BOX, PLACEHOLDER

report = {"fichier": PPTX.name, "format_pouces": [SW, SH], "ratio": round(SW / SH, 4),
          "diapositives": len(prs.slides._sldIdLst), "problemes": [], "avertissements": [],
          "notes_manquantes": [], "alt_manquant": [], "debordements": [], "hors_zone": [],
          "polices": set(), "sans_texte_alt_informatif": 0}

for i, slide in enumerate(prs.slides, 1):
    if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
        report["notes_manquantes"].append(i)
    for sh in slide.shapes:
        descr = sh._element._nvXxPr.cNvPr.get("descr")
        if not descr:
            report["alt_manquant"].append({"slide": i, "forme": sh.shape_type,
                                           "texte": (sh.text_frame.text[:40]
                                                     if sh.has_text_frame else "")})
        l, t = sh.left / EMU_IN, sh.top / EMU_IN
        w, h = sh.width / EMU_IN, sh.height / EMU_IN
        if l < -0.01 or t < -0.01 or l + w > SW + 0.01 or t + h > SH + 0.01:
            report["hors_zone"].append({"slide": i, "forme": str(sh.shape_type),
                                        "zone": [round(l, 2), round(t, 2), round(w, 2), round(h, 2)]})
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                if r.font.name:
                    report["polices"].add(r.font.name)
        text = tf.text
        if not text.strip():
            continue
        sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
        size = max(sizes) if sizes else 18
        bold = any(r.font.bold for p in tf.paragraphs for r in p.runs)
        avail_w = w * 72 - 6
        lines = 0
        for para in text.split("\n"):
            lines += n_lines(para, size, avail_w, bool(bold)) if para.strip() else 1
        needed = lines * size * 1.14
        if needed > h * 72 + 2:
            report["debordements"].append({
                "slide": i, "texte": text[:48].replace("\n", " / "),
                "hauteur_dispo_pt": round(h * 72, 1), "hauteur_estimee_pt": round(needed, 1),
                "taille_pt": size, "lignes": lines})

report["polices"] = sorted(report["polices"])
report["liens_externes"] = sum(
    1 for slide in prs.slides for sh in slide.shapes
    if sh.has_text_frame and ("http" in sh.text_frame.text))
report["total_formes"] = sum(len(s.shapes) for s in prs.slides)
report["total_notes_caracteres"] = sum(
    len(s.notes_slide.notes_text_frame.text) for s in prs.slides if s.has_notes_slide)

ok = (len(prs.slides._sldIdLst) == 84 and not report["notes_manquantes"]
      and not report["alt_manquant"] and not report["debordements"]
      and not report["hors_zone"] and report["ratio"] > 1.77)
report["verdict"] = "CONFORME" if ok else "A CORRIGER"
slim={k:v for k,v in report.items() if k not in ("alt_manquant","debordements")}
slim["nb_alt_manquant"]=len(report["alt_manquant"]);slim["exemples_alt_manquant"]=report["alt_manquant"][:5]
slim["nb_debordements"]=len(report["debordements"]);slim["debordements"]=report["debordements"]
slim["notes_manquantes"]=report["notes_manquantes"];slim["hors_zone"]=report["hors_zone"]
print(json.dumps(slim, ensure_ascii=False, indent=1))
print("\nVERDICT:", report["verdict"])
