# Strumenti di assemblaggio

Script usati per Volume 2 · Sezione 1, da adattare per ogni nuovo capitolo. Le regole di formattazione fisse sono in [`../ISTRUZIONI_ASSEMBLAGGIO.md`](../ISTRUZIONI_ASSEMBLAGGIO.md) — questa cartella contiene il codice che le applica.

## Ordine di esecuzione

1. **`assembla_capitolo.py`** — prende il `.docx` grezzo del capitolo e produce il `.docx` finale (font, box, grassetti, interruzioni di pagina, fix strutturali). Cercare i blocchi `# --- MODIFICARE QUI` e aggiornarli con i dati del nuovo capitolo prima di eseguire.
2. **`docx_to_md.py`** — dal `.docx` finale produce il `.md` leggibile.
3. Per il PDF: convertire il `.docx` finale con LibreOffice (`soffice --headless --convert-to pdf file.docx`), controllare pagine bianche e titoli tagliati.
4. Per l'ePub:
   - `pandoc file.docx -o file.epub --extract-media=media`
   - `estrai_colori_riquadri.py` (genera `riquadro_colors.json`)
   - `epub_postprocess.py` (rimuove i QR, aggiunge i link veri, riapplica i colori — aggiornare i blocchi `MODIFICARE QUI` con i riferimenti del nuovo capitolo)
   - Ricompattare rispettando le regole ePub (`mimetype` primo file, non compresso — vedi ISTRUZIONI_ASSEMBLAGGIO.md §7)

## `manuale-reference.docx`

Reference-doc di Pandoc con gli stili della collana già impostati: **Normal** (corpo del testo) in Georgia 11pt/interlinea 1,15; **Heading 1-3** e **Title** in Gill Sans grassetto, senza nessuna evidenziazione residua; **Block Text** (il target dei blockquote `>` in markdown) come riquadro — Gill Sans, sfondo chiaro `F3EEE2`, nessun bordo.

Va allegato/richiamato quando si genera un `.docx` a partire da markdown, così la formattazione di base è già coerente con la collana prima ancora di passare da `assembla_capitolo.py`:

```
pandoc capitolo.md -o capitolo.docx --reference-doc=strumenti/manuale-reference.docx
```

Non sostituisce `assembla_capitolo.py` (che gestisce box, QR, grassetti mirati, correzioni Google Docs, ecc.) ma dà un punto di partenza già pulito, specialmente utile se il capitolo nasce direttamente in markdown invece che in Google Docs.

Per rigenerarlo o aggiornarlo (es. se cambiano i colori/font della collana): partire da `pandoc -o base.docx --print-default-data-file reference.docx` e modificare gli stili con `python-docx`, avendo cura di rimuovere sempre i riferimenti `*Theme` da `w:rFonts` (altrimenti il tema vince sul font esplicito) e di rispettare l'ordine degli elementi in `w:pPr`/`w:rPr` (vedi ISTRUZIONI_ASSEMBLAGGIO.md §6).

## Nota sui percorsi

Tutti gli script hanno in cima una variabile `SCRATCH`/`SRC`/`OUT` con percorsi assoluti della sessione in cui sono stati scritti — vanno sempre aggiornati a inizio lavoro, sono lì solo perché lo script sia eseguibile as-is durante lo sviluppo, non perché siano fissi.
