# Raccolta URL — manuale/volume3

File di lavoro per accumulare gli indirizzi di artefatti e video di Volume 3 man mano che vengono prodotti, così da averli pronti al momento di assemblare il testo finale (stessa logica di `links.json` usato per Volume 2 · Sezione 1, ma qui **versionato nel repo** invece che in uno scratchpad di sessione — altrimenti si perde da una conversazione all'altra).

## Struttura delle cartelle

Mirror esatto di `volume2`, una sotto-cartella per sezione:

```
volume3/
├── links.json                        → questo file di raccolta URL (unico per tutto il volume)
├── README.md
└── sezioni/{M}/
    ├── img/                          → immagini/cartine della sezione
    └── artefatti/                    → pagine HTML interattive (01-nome.html, 02-...)
```

## `links.json`

Un unico oggetto piatto: `"id": "url"`.

```json
{
  "art1": "https://baldassarrefrancesco70.github.io/manuale/volume3/sezioni/1/artefatti/01-populismo-migrazioni.html",
  "video1": "https://drive.google.com/file/d/FILEID/view"
}
```

**Convenzione degli id**

- `art1`, `art2`, `art3`, ... per gli artefatti — numerati nell'ordine in cui compaiono nel testo (non deve necessariamente coincidere con il numero nel nome del file in `sezioni/{M}/artefatti/`, vedi `ISTRUZIONI_ASSEMBLAGGIO.md`).
- `video1`, `video2`, `video3`, ... per i video, stessa logica di numerazione per ordine di comparsa.

**Formato degli URL**

- Artefatti: indirizzo pubblicato su GitHub Pages, es. `https://baldassarrefrancesco70.github.io/manuale/volume3/sezioni/1/artefatti/NN-slug.html`.
- Video: link di condivisione Google Drive, es. `https://drive.google.com/file/d/FILEID/view`.

## Come aggiornarlo

Ogni volta che un nuovo artefatto viene pubblicato nel repo o un nuovo video viene caricato su Drive, aggiungere una riga a `links.json` con il prossimo id libero e il suo URL. Non serve altro — titoli e descrizioni per i box e l'indice finale si scrivono al momento dell'assemblaggio (vedi `DESC`/`TITLES` in `strumenti/assembla_capitolo.py`).
