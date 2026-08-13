# manuale/volume2

Materiali pubblicati (immagini, artefatti) per il Volume 2 del manuale.

## Struttura delle cartelle

```
volume2/
└── sezioni/{M}/
    ├── img/                          → immagini/cartine/infografiche della sezione
    └── artefatti/                    → pagine HTML interattive (01-nome.html, 02-...)
```

Le pagine GitBook generate da `strumenti/md_to_gitbook.py` (dentro `Sezione 1` del volume) vivono invece in `gitbook/volume2/`, non qui: questa cartella contiene solo i materiali grezzi (immagini e artefatti), referenziati sia dal PDF (via `strumenti/md_to_tex.py`) sia dalle pagine GitBook.

## Sezione 1 — Come nasce l'economia del mondo moderno

Vedi `sezioni/1/README.md` per l'elenco di tutte le immagini e gli artefatti già pubblicati. I 23 video della sezione restano su Google Drive (link nel `README.md` alla radice del repo).

## Sezione 2 — Come nasce la politica del mondo moderno

Vedi `sezioni/2/README.md` per l'elenco dei 24 artefatti e delle 29 immagini (6 cartine + 23 infografiche) già pubblicati. Le infografiche sono caricate come file grezzi ma non ancora wired nella configurazione della sezione (`strumenti/sezioni/volume2_sezione2_politica.py`): 24 marcatori nel `.md` contro 23 file, richiede verifica umana prima di popolare `INFOGRAFICA_FILES`. I video restano su YouTube (link in `tabella-conversione-video-volume2.csv`, non replicati qui).
