# =============================================================================
# Estrae il colore di sfondo originale di ogni riquadro/box dal .docx
# assemblato, per poterlo riapplicare nell'ePub (pandoc non lo porta con sé).
# Da eseguire prima di epub_postprocess.py.
# =============================================================================
import docx, json
from docx.oxml.ns import qn

# --- MODIFICARE QUI: percorsi per il capitolo corrente ----------------------
SCRATCH = "/private/tmp/claude-502/-Users-francescobaldassarre/2e72ca84-06e7-43ac-afe9-c15a06aa7354/scratchpad"
SRC = f"{SCRATCH}/volume2_sezione1_assemblato.docx"

d = docx.Document(SRC)


def get_text(el):
    return "".join(t.text or "" for t in el.findall('.//' + qn('w:t')))


colors = {}
for t in d.tables:
    full = get_text(t._tbl).strip()
    if not full:
        continue
    cell = t.rows[0].cells[0]
    tcPr = cell._tc.find(qn('w:tcPr'))
    if tcPr is None:
        continue
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        continue
    fill = shd.get(qn('w:fill'))
    if fill and fill != "auto":
        colors[full[:60]] = fill

json.dump(colors, open(f"{SCRATCH}/riquadro_colors.json", "w"), ensure_ascii=False, indent=1)
print("estratti", len(colors), "colori ->", f"{SCRATCH}/riquadro_colors.json")
print("\nNOTA: in epub_postprocess.py il dizionario DISTINCT va aggiornato a mano")
print("con i prefissi di testo di questo nuovo capitolo (i box F3EEE2 sono")
print("gestiti automaticamente, non serve elencarli).")
