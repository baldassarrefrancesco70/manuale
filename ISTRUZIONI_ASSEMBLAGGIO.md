# Istruzioni per l'assemblaggio dei capitoli — Manuale

Riepilogo del processo usato per Volume 2 · Sezione 1, da riapplicare tale e quale ai prossimi capitoli/sezioni.

## 1. Materiali e repository

Struttura del repo (`baldassarrefrancesco70/manuale`, pubblicato via GitHub Pages su `https://baldassarrefrancesco70.github.io/manuale/`):

```
manuale/
├── img/                              → icone condivise (video.png, risorse_digitali.png, atlante.png)
└── volume{N}/sezioni/{M}/
    ├── img/                          → immagini/cartine della sezione
    ├── artefatti/                    → pagine HTML interattive (01-nome.html, 02-...)
    └── README.md                     → indice dei link della sezione
```

- **Icone**: `video.png` (clapperboard) per i video; `risorse_digitali.png` (lampadina/puzzle) per gli artefatti/risorse digitali; `atlante.png` (bussola, creata su misura) solo per i rimandi interni alle cartine.
- **Video**: caricati su Google Drive, link nel formato `https://drive.google.com/file/d/FILEID/view`.
- **QR code**: generati con la libreria python `qrcode`, uno per ciascun URL (risorsa o video).

## 2. Testo di partenza

Il documento arriva come `.docx` con due parti:
1. Il testo del capitolo/sezione (con box colorati preesistenti: Apertura/Chiusura di sezione, Prerequisiti, Focus, Nodi storiografici, Per approfondire, Atlante).
2. Un documento "mappa dei segnaposto" separato che indica dove inserire artefatti/video/cartine nel testo.

**Attenzione**: i box colorati del documento originale sono spesso **tabelle Word**, non semplici paragrafi — `python-docx`'s `document.paragraphs` **non le vede**. Bisogna sempre camminare `document.element.body` includendo sia `<w:p>` che `<w:tbl>` per avere la mappa reale della struttura (altrimenti si rischia di credere vuote sezioni che in realtà sono dentro tabelle).

Prima di costruire la mappa di inserimento, verificare sempre:
- quali segnaposto della mappa corrispondono davvero a un titolo/sezione nel testo (i nomi possono essere cambiati nel frattempo);
- se ci sono artefatti/video extra non presenti nella mappa (è già successo: 3 artefatti "in più" trovati solo guardando i file nel repo);
- se ci sono sezioni della mappa il cui testo sembra mancante — controllare con `pandoc -t markdown` oltre a `python-docx`, perché pandoc individua meglio contenuto dentro tabelle.

## 3. Regole di formattazione (fissate dopo iterazione con l'autore)

**Font e dimensioni**
- Corpo del testo (paragrafi numerati, es. "1.1 ..."): **Georgia**, 11pt
- Riquadri originali (Apertura/Chiusura, Prerequisiti, Focus, Per approfondire, barre "Capitolo N", Atlante): **Gill Sans**, 11pt
- Box risorse/video (creati da noi): **Avenir Next** — titolo 11,5pt, descrizione 9,5pt, indirizzo 8,5pt

**Spaziatura**
- Interlinea 1,15 ovunque (corpo e riquadri)
- Spazio dopo paragrafo: **2pt** (prima: 0pt)
- Intorno ai box risorse/video: ~12pt di margine prima e dopo, rispetto al testo circostante

**Box risorse/video**
- Nessun bordo, sfondo leggero (`F3EEE2`)
- QR code a sinistra (0,5in nei box inline, 0,4in nell'indice finale), centrato verticalmente
- A destra: icona + titolo in grassetto sulla stessa riga, poi (nell'indice finale) una breve descrizione, poi l'indirizzo per esteso in chiaro
- Nessuna intestazione numerata tipo "Video 3" — solo icona e titolo

**Grassetto**
- Il corpo del testo originale ha già grassetti; nei riquadri (che non ne avevano) aggiungerne un po' a mano sui concetti/nomi/date chiave, senza esagerare — 3-8 per riquadro è una buona misura

**Capitoli**
- Ogni "Capitolo N" inizia sempre su una pagina nuova (vedi §5 per come farlo senza creare pagine bianche)
- Per il resto: **flusso naturale**, nessun automatismo per "tenere insieme" righe/titoli orfani — si è provato più volte e ha sempre creato più problemi di quanti ne risolvesse. Meglio controllare a mano e correggere i casi singoli nel `.docx` finale.

## 4. Link — regola definitiva

**Nessun collegamento ipertestuale cliccabile nel `.docx`**, né interno né esterno. Motivo: costruire `<w:hyperlink>` a mano in OOXML è fragile (l'ordine degli elementi XML dentro `rPr`/`pPr`/`tcPr` è vincolato dallo schema; se sbagliato, Word "ripara" il file in silenzio e butta via il link). Ci sono voluti diversi tentativi falliti prima di arrivare a questa regola.

- **Box risorse/video**: indirizzo per esteso in chiaro (funziona comunque: molti visualizzatori PDF lo riconoscono e lo rendono cliccabile a video, e resta comunque copiabile; il QR code è il modo pensato per l'accesso rapido).
- **Rimandi alle cartine** ("Atlante · cartina N · ... — vedi in fondo al volume"): testo semplice, nessun link, nel `.docx`.

**Nell'ePub è diverso**: HTML/XHTML supporta `<a href="#id">` in modo semplice e affidabile. Lì sì, si possono (anzi si devono) fare i link veri — vedi §6.

## 5. Interruzione di pagina per i capitoli (senza pagine bianche)

Usare **`pageBreakBefore`** a livello di paragrafo (non `<w:br w:type="page"/>` inserito a mano in mezzo al testo — quello ha causato pagine bianche più volte).

```python
def set_page_break_before(tbl_el):
    first_p = tbl_el.find(qn('w:tr')).find(qn('w:tc')).find(qn('w:p'))
    pPr = first_p.find(qn('w:pPr')) or OxmlElement('w:pPr')  # inserire se assente
    pbb = OxmlElement('w:pageBreakBefore')
    insert_in_pPr_order(pPr, pbb, 'w:pageBreakBefore')  # rispettare l'ordine CT_PPr
```

**Prima di applicarlo**, rimuovere eventuali paragrafi vuoti residui del documento originale subito prima della tabella "Capitolo N" (retaggio di vecchi page-break manuali) — altrimenti quei paragrafi vuoti + il nuovo page-break creano comunque una pagina bianca.

## 6. Problemi tecnici incontrati e soluzioni (da NON rifare)

1. **Evidenziazione gialla "fantasma"**: non era un `<w:highlight>` sui singoli run, ma **nella definizione dello stile "Heading3"** in `styles.xml` (molti paragrafi di corpo testo erano stati erroneamente marcati con questo stile). Va rimossa lì, una volta sola.
2. **Grassetto che "smangia" la formattazione circostante**: quando lo stile ha bold=1 di default e il run lo sovrascrive con `w:b w:val="0"`, bisogna **impostare esplicitamente `val="0"`**, mai limitarsi a rimuovere l'elemento `<w:b>` (altrimenti il testo eredita bold=1 dallo stile).
3. **Bordi/ombreggiature che rompono il file**: gli elementi dentro `<w:tcPr>` (tcBorders, shd, vAlign...) e dentro `<w:pPr>` (keepNext, shd, spacing...) hanno un **ordine fisso nello schema OOXML**. Un `.append()` ingenuo li mette fuori ordine: il file sembra funzionare ma Word lo "ripara" silenziosamente, con effetti collaterali imprevedibili (compresi i link che spariscono). Usare sempre una funzione che inserisce nella posizione giusta rispettando l'ordine CT_TcPr / CT_PPr.
4. **ID di bookmark duplicati**: il documento originale (esportato da Google Docs) ne aveva parecchi. Non causano danni visibili ma fanno fallire la validazione formale — meglio deduplicarli.
5. **Margini di pagina con attributi mancanti** (`w:gutter`): altro retaggio dell'export Google Docs, va aggiunto per la validità dello schema.
6. **Verificare sempre con uno script di validazione** prima di considerare un `.docx` finito (skill `docx`, `scripts/office/validate.py`) — intercetta questi problemi prima che arrivino all'utente.

## 7. Processo ePub

1. Convertire il `.docx` finale con **pandoc**: `pandoc file.docx -o file.epub --extract-media=media`
2. Estrarre l'archivio ePub e post-processare `EPUB/text/ch001.xhtml`:
   - Aggiungere `id="cartinaN"` ai quattro titoli delle cartine nell'Atlante
   - Sostituire "vedi in fondo al volume" con un vero `<a href="#cartinaN">`
   - Rimuovere le immagini dei QR code (riconoscibili dallo stile `width:0.4in`/`0.5in`) e collassare la colonna della tabella che li conteneva (altrimenti resta uno spazio vuoto a sinistra del box)
   - Trasformare gli indirizzi in chiaro (`<em>https://...</em>`) in `<a href="...">` veri
3. Ricompattare l'archivio rispettando le regole del formato ePub: `mimetype` deve essere il **primo file**, salvato **non compresso** (`zip -X -q out.epub mimetype`, poi `zip -X -rq out.epub . -x mimetype`)
4. Verificare (idealmente con `epubcheck`, se disponibile senza dover compilare da sorgente — l'installazione via Homebrew può risultare molto lenta) o quantomeno controllare la buona formazione XML con `lxml` e ispezionare a occhio in un browser puntato su un server locale (`python3 -m http.server`).

## 8. Ordine di lavoro consigliato per il prossimo capitolo

1. Caricare testo + mappa segnaposto, leggere l'intera struttura (paragrafi **e** tabelle)
2. Ricostruire la mappa di inserimento (artefatti/video/cartine), segnalando subito eventuali discrepanze prima di scrivere codice
3. Applicare le regole di formattazione fisse di questo file (font, spaziatura, box, niente link nel docx, interruzione pagina per capitolo)
4. Validare lo schema del `.docx` prima di generare il PDF
5. Generare PDF di verifica, controllare pagine bianche/titoli tagliati a occhio
6. Solo dopo l'ok sul `.docx`, generare l'ePub e applicare i link reali + rimozione QR
