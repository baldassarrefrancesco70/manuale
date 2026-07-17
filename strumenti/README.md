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

## Nota sui percorsi

Tutti gli script hanno in cima una variabile `SCRATCH`/`SRC`/`OUT` con percorsi assoluti della sessione in cui sono stati scritti — vanno sempre aggiornati a inizio lavoro, sono lì solo perché lo script sia eseguibile as-is durante lo sviluppo, non perché siano fissi.
