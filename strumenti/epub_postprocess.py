# =============================================================================
# Post-elaborazione ePub: rimuove i QR, trasforma URL e rimandi cartine in
# link cliccabili veri, ripristina i colori di sfondo (persi da pandoc).
# Da eseguire dopo: pandoc capitolo.docx -o capitolo.epub --extract-media=media
# ed estrazione dell'archivio. Vedi ISTRUZIONI_ASSEMBLAGGIO.md §7.
# =============================================================================
import re, json
from lxml import etree

# --- MODIFICARE QUI: percorso del file xhtml del nuovo capitolo ------------
SCRATCH = "/private/tmp/claude-502/-Users-francescobaldassarre/2e72ca84-06e7-43ac-afe9-c15a06aa7354/scratchpad"
F = f"{SCRATCH}/epub_work/extracted/EPUB/text/ch001.xhtml"

# ---------------------------------------------------------------------------
# Step A: regex-based fixes (QR removal, link-ification) -- same as before,
# reapplied to the freshly regenerated file.
# ---------------------------------------------------------------------------
html = open(F, encoding="utf-8").read()

# --- MODIFICARE QUI: titoli delle cartine del nuovo capitolo -> ancore -----
MAP_IDS = {
    "Recinzione dei campi comuni, 1700-1800": "cartina1",
    "Diffusione della rivoluzione industriale in Europa": "cartina2",
    "Lo sviluppo industriale del Regno Unito": "cartina3",
    "Le rotte del commercio atlantico, XVI-XIX secolo": "cartina4",
}
for title, anchor_id in MAP_IDS.items():
    old = f"<p><strong>{title}</strong></p>"
    new = f'<p id="{anchor_id}"><strong>{title}</strong></p>'
    count = html.count(old)
    assert count == 1, f"expected 1 occurrence of heading {title!r}, found {count}"
    html = html.replace(old, new, 1)

pattern = re.compile(r"(Atlante · cartina (\d) · [^—]+— )vedi in fondo al volume")
def repl(m):
    prefix, num = m.group(1), m.group(2)
    return f'{prefix}<a href="#cartina{num}">vedi in fondo al volume</a>'
html, n_cartina = pattern.subn(repl, html)
print("cartina cross-reference links added:", n_cartina)

html, n_qr = re.subn(
    r'<img src="[^"]*" style="width:0\.[45]in;height:0\.[45]in" alt="" />',
    "",
    html,
)
print("QR images removed:", n_qr)

html, n_url = re.subn(
    r"<em>(https?://[^<]+)</em>",
    r'<em><a href="\1">\1</a></em>',
    html,
)
print("plain URLs linkified:", n_url)

collapse_pattern = re.compile(
    r'<colgroup>\n<col style="width: \d+%" />\n<col style="width: \d+%" />\n</colgroup>\n'
    r'<thead>\n<tr>\n<th style="text-align: center;"></th>\n'
    r'(<th style="text-align: left;">.*?</th>)\n</tr>\n</thead>',
    re.S,
)
html, n_collapse = collapse_pattern.subn(
    '<colgroup>\n<col style="width: 100%" />\n</colgroup>\n<thead>\n<tr>\n\\1\n</tr>\n</thead>',
    html,
)
print("collapsed empty QR column in", n_collapse, "tables")

open(F, "w", encoding="utf-8").write(html)

# ---------------------------------------------------------------------------
# Step B: restore background colors on the riquadri / resource boxes, using
# the color map extracted from the source docx (riquadro_colors.json).
# ---------------------------------------------------------------------------
colors = json.load(open(f"{SCRATCH}/riquadro_colors.json", encoding="utf-8"))
# --- MODIFICARE QUI: un prefisso di testo breve e univoco per ciascun
#     riquadro del nuovo capitolo (i colori F3EEE2 dei box risorse/video
#     sono già gestiti automaticamente più sotto, non serve elencarli qui) --
DISTINCT = {
    "SEZIONE 1": "ffca2e",
    "APERTURA DI SEZIONE": "f4cccc",
    "Capitolo 1 -": "fff2cc",
    "Prerequisiti": "d9ead3",
    "Focus - Adam Smith": "fff2cc",
    "Nodi storiografici controversi": "cc0000",
    "Per approfondireCarlo M. Cipolla, Storia economica dell'Euro": "cfe2f3",
    "Capitolo 2 -": "fff2cc",
    "Il filo del discorsoUn mondo con più braccia": "d9ead3",
    "Focus - La meccanizzazione della filatura": "fff2cc",
    "Focus - La parola": "fff2cc",
    "Focus - L'organizzazione del movimento operaio": "fff2cc",
    "Per approfondirePer un quadro d'insieme": "cfe2f3",
    "Capitolo 3 -": "ffe599",
    "Il filo del discorsoAlla metà dell'Ottocento": "d9ead3",
    "Focus - Il sorpasso": "fff2cc",
    "Focus - I templi del consumo": "fff2cc",
    "Per approfondireSul dibattito relativo ai consumi": "cfe2f3",
    "CHIUSURA DI SEZIONE": "f4cccc",
    "ATLANTE": "38761d",
    "RISORSE DIGITALI": "1155cc",
}

NSMAP = {"x": "http://www.w3.org/1999/xhtml"}
parser = etree.XMLParser(remove_blank_text=False)
tree = etree.parse(F, parser)
root = tree.getroot()

def local_text(el):
    return "".join(el.itertext())

tables = root.findall(".//x:table", namespaces=NSMAP)
n_distinct = 0
n_box = 0
for tbl in tables:
    text = local_text(tbl).strip()
    ths = tbl.findall(".//x:th", namespaces=NSMAP)
    matched = False
    for prefix, hexcolor in DISTINCT.items():
        if text.startswith(prefix):
            for th in ths:
                existing = th.get("style", "")
                th.set("style", (existing + f";background-color:#{hexcolor}").lstrip(";"))
            n_distinct += 1
            matched = True
            break
    if matched:
        continue
    # resource/video box: has a link to github.io or drive.google.com
    links = tbl.findall(".//x:a", namespaces=NSMAP)
    if any((lk.get("href") or "").startswith("http") for lk in links):
        for th in ths:
            existing = th.get("style", "")
            th.set("style", (existing + ";background-color:#F3EEE2").lstrip(";"))
        n_box += 1

print("distinct-colored riquadri matched:", n_distinct)
print("resource/video boxes colored:", n_box)

tree.write(F, xml_declaration=True, encoding="UTF-8", standalone=None)
print("done")
