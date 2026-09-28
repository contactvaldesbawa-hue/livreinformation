"""Build original, accessible SVG planning assets; not yet embedded in a book."""
from pathlib import Path
from html import escape
P=Path(__file__).parent
NAVY='#101a35'; CYAN='#27c9d8'; PURPLE='#8668ca'

def diagram(name, rows):
    width=1400; height=130+len(rows)*144
    s=[f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" role="img" aria-labelledby="title desc">',f'<title id="title">{escape(name.replace("_"," "))}</title><desc id="desc">'+escape(' ; '.join(' puis '.join(r) for r in rows))+'</desc>',f'<rect width="{width}" height="{height}" rx="25" fill="#f5f8fc"/>']
    for j,row in enumerate(rows):
        n=len(row); gap=18; cw=min(243,(width-96-(n-1)*gap)/n); left=(width-(n*cw+(n-1)*gap))/2; y=70+j*144
        for i,label in enumerate(row):
            x=left+i*(cw+gap)
            s.append(f'<rect x="{x:.1f}" y="{y}" width="{cw:.1f}" height="80" rx="15" fill="{NAVY if i%2==0 else PURPLE}"/>')
            # two lines only for long text, with explicitly controlled wrapping
            words=label.split(); lines=['']
            for word in words:
                if len(lines[-1])+len(word)+1>19 and lines[-1]: lines.append('')
                lines[-1]=(lines[-1]+' '+word).strip()
            lines=lines[:2]
            for k,line in enumerate(lines):
                s.append(f'<text x="{x+cw/2:.1f}" y="{y+37+24*k-(12 if len(lines)==2 else 0)}" text-anchor="middle" font-family="DejaVu Sans, Arial, sans-serif" font-size="17" font-weight="bold" fill="white">{escape(line)}</text>')
            if i<n-1:
                ax=x+cw+3; ay=y+40
                s.append(f'<path d="M {ax:.1f} {ay} h 13 m -5 -5 l 5 5 -5 5" fill="none" stroke="{CYAN}" stroke-width="3"/>')
        if j<len(rows)-1: s.append(f'<path d="M 700 {y+87} v 43" stroke="{CYAN}" stroke-width="3"/>')
    s.append('</svg>')
    (P/name).write_text(''.join(s),encoding='utf-8')

diagram('systeme_createur.svg', [['Compétence','Niche','Audience','Contenu','Confiance'],['Offre utile','Revenus possibles','Système documenté']])
diagram('delegation_ia.svg', [['Humain','Texte ou voix','Brief structuré','IA / agent'],['Recherche et organisation','Contrôle humain','Publication validée']])
diagram('production_video.svg', [['Recherche','Idée','Brief','Script'],['Visuels et voix','Montage','Publication','Mesure']])
diagram('mesures.svg', [['Portée','Intérêt','Demandes','Ventes encaissées']])
print('SVG créés:',*[p.name for p in P.glob('*.svg')])
