# WebArch — EASYONE MINIMAL

## Stato del progetto

**Alpha congelata: EASYONE MINIMAL 1.0.3 + HOME 1.1.**

Il 14 settembre 2026 il responsabile del progetto ha confermato: «funziona tutto come previsto», richiedendo il congelamento dell'alpha e l'avvio della beta. La conferma riguarda il suo utilizzo in Excel; non viene trasformata in una certificazione automatica di tutti i casi di test.

La beta B0 parte dagli stessi componenti produttivi: nessuna riscrittura, nessuna rimozione del debug e nessun ritorno alla macchina transazionale. La versione interna VBA resta EASYONE-MINIMAL-1.0.3; HOME resta 1.1.

## Disponibilità dei file: distinzione importante

Questo aggiornamento del repository pubblica il **registro del congelamento, gli hash dei 36 componenti, il piano beta e il verificatore dell'archivio**. Il pacchetto completo è stato preparato e consegnato separatamente nella conversazione di progetto:

`WebArch_ALPHA_1_0_3_HOME_1_1_CONGELATA.zip`

SHA-256: `e94986fbee71cfcb380fdb53503a91058c812efcc3cbabac03e8c800872ef6f0`

**L'archivio completo, il template e tutti i sorgenti NON sono ancora caricati in questo Git tree o come asset di una release GitHub. Scaricare il repository non equivale a scaricare WebArch completo.** La pubblicazione integrale è un'attività aperta della beta. Il file .xlsm operativo con i dati dell'utente non è stato caricato e non deve essere pubblicato su questo repository pubblico.

I file `BUILD/`, `BUILD_WEBARCH.cmd` e `PATCH_NOTES_v1.1.1.md` appartengono al builder storico pre-MINIMAL: sono conservati per tracciabilità, **non sono la procedura di installazione della beta**.

## Baseline conservata

- Aggiorna tutto: importazione, validazione, deduplicazione, scrittura dati, analitiche, dashboard e report; salvataggio finale.
- Stati protocollo presenti ma non mappati: WARNING registrato, protocollo conservato e aggiornamento non bloccato per questo motivo.
- Correzioni PDF/XLS, accodamento staging, confronto dei testi simili a date e anteprima PA3.
- Debug del salvataggio mantenuto, compreso DisplayAlerts al salvataggio finale.
- Esportazione ordinaria come bozza tecnica; nessuna autorizzazione Gold implicita.
- HOME 1.1: collegamenti interni ai fogli, senza dipendenza dal nome del workbook.

## Documenti correnti

- [Congelamento e perimetro](releases/alpha-1.0.3-home1.1/FREEZE.md)
- [Identità dell'archivio](releases/alpha-1.0.3-home1.1/artifact.json)
- [Hash dei componenti produttivi](releases/alpha-1.0.3-home1.1/COMPONENTS.sha256)
- [Piano beta B0](docs/BETA_B0.md)
- [Verifica dell'archivio ricevuto](tools/verify_frozen_archive.py)

Non reinstallare o reimportare moduli nella copia funzionante soltanto per cambiare fase del progetto. Conservare l'alpha separatamente e svolgere i test beta su copie prive di dati da pubblicare.
