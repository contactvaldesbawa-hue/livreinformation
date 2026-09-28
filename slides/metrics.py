"""Estimation déterministe de la largeur du texte.

Le rendu réel de PowerPoint n'est pas disponible dans cet environnement
(aucun libreoffice/soffice). Pour éviter les débordements, la mise en page
est calculée à partir des métriques d'une police libre proche d'Arial
(DejaVu Sans) : largeur moyenne par caractère, ajustée par famille.
"""
from PIL import ImageFont
from pathlib import Path

_DEJAVU = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
_cache = {}


def _font(bold: bool, px: int):
    key = (bold, px)
    if key not in _cache:
        path = _BOLD if bold else _DEJAVU
        if Path(path).exists():
            _cache[key] = ImageFont.truetype(path, px)
        else:  # repli : facteur constant
            _cache[key] = None
    return _cache[key]


# Arial est environ 5 % plus étroite que DejaVu Sans
_ARIAL_FACTOR = 0.95


def text_width_pt(text: str, size_pt: float, bold: bool = False) -> float:
    f = _font(bold, 100)
    if f is None:
        return len(text) * size_pt * 0.52
    return f.getlength(text) / 100.0 * size_pt * _ARIAL_FACTOR


def wrap_lines(text: str, size_pt: float, width_pt: float, bold: bool = False) -> list[str]:
    words = text.split()
    lines, cur = [], ""
    for w in words:
        cand = f"{cur} {w}".strip()
        if cur and text_width_pt(cand, size_pt, bold) > width_pt:
            lines.append(cur)
            cur = w
        else:
            cur = cand
    if cur:
        lines.append(cur)
    return lines


def n_lines(text: str, size_pt: float, width_pt: float, bold: bool = False) -> int:
    return max(1, len(wrap_lines(text, size_pt, width_pt, bold)))


def fit_size(text: str, width_pt: float, height_pt: float, start: float,
             min_size: float = 11.0, bold: bool = False, line_spacing: float = 1.18) -> float:
    """Réduit la taille de police jusqu'à ce que le texte tienne dans la zone."""
    size = start
    while size > min_size:
        lines = n_lines(text, size, width_pt, bold)
        if lines * size * line_spacing <= height_pt:
            return size
        size -= 0.5
    return min_size
