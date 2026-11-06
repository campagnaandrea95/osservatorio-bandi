# Osservatorio bandi

Avvisi pubblici per contributi a imprese, startup e ricerca: europei, nazionali e regionali.

Sito statico pubblicato con GitHub Pages. I file in `data/` vengono aggiornati automaticamente:

- `data/opportunita.json` – opportunità aperte, in arrivo e scadute da meno di 90 giorni
- `data/schede/<id>.json` – schede di approfondimento
- `data/schede/index.json` – elenco delle schede disponibili
- `data/meta.json` – data dell'ultimo aggiornamento

Le schede si richiedono dal pulsante "Richiedi scheda" (modulo Google): le richieste vengono evase ogni 4 ore circa. Il pulsante "Scarica Word" crea il documento .docx della scheda direttamente nel browser (libreria `docx`, licenza MIT, in `lib/`).

L'elenco è informativo e non sostituisce l'avviso ufficiale.
