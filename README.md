# WebArch — EASYONE MINIMAL

## Stato del progetto

**Alpha congelata: EASYONE MINIMAL 1.0.3 + HOME 1.1.**

Il 14 settembre 2026 il responsabile del progetto ha confermato: «funziona tutto come previsto», richiedendo il congelamento dell'alpha e l'avvio della beta. La conferma riguarda il suo utilizzo in Excel; non viene trasformata in una certificazione automatica di tutti i casi di test.

La beta parte dagli stessi componenti produttivi dell'Alpha e resta strettamente additiva: nessuna riscrittura dell'Alpha, nessuna rimozione del debug e nessun ritorno alla macchina transazionale. La versione interna VBA dell'Alpha resta EASYONE-MINIMAL-1.0.3; HOME resta 1.1.

## Beta corrente

**M10.4 / M11.4 — System Integration Hardening — gate candidate, RC1 non congelata.**

M10.4 supera la verifica dei singoli milestone e aggiunge la protezione sistemica contro lost-update tra `OP_PROTOCOLLI` (M3) e `OP_PROTOCOLLI_PLUS` (M4): entrambe le viste verificano lo stato M2 applicato e le bozze `PENDING` prima di qualsiasi commit. Restano inoltre attivi gli hardening M4/M5/M7/M8 e i gate fail-closed di M10.

Evidenza offline corrente:

- full-project audit: **1082/1082 PASS**;
- package release verifier: **67/67 PASS**;
- M10 models: **59/59 PASS**;
- functional-flow models: **31/31 PASS**;
- system-integration models: **13/13 PASS**;
- random integration: **5.000 iterazioni / 20.000 controlli PASS**;
- integration mutations: **7/7 PASS**;
- shared-state audit completo: **18/18 PASS**;
- package suites: **7/7 PASS**;
- manifest degli ZIP estratti: **42/42 PASS** per M10.4, M11.4 e consegna combinata.

Microsoft VBE/Excel reale non è ancora stato eseguito per questa candidata: compilazione VBAProject, eventi Workbook, ListObject runtime, Shape/OnAction, Save/reopen, PDF e performance restano il gate necessario prima di RC1.

## Disponibilità dei file

Il repository pubblica il registro del congelamento Alpha e i record delle release Beta con stato, QA e checksum. La consegna completa M10.4/M11.4 con sorgenti e ZIP esatti è archiviata nella Project Library del progetto WebArch; il file `.xlsm` operativo con dati dell'utente non è pubblicato e non deve essere caricato in questo repository pubblico.

I file `BUILD/`, `BUILD_WEBARCH.cmd` e `PATCH_NOTES_v1.1.1.md` appartengono al builder storico pre-MINIMAL: sono conservati per tracciabilità, **non sono la procedura di installazione della beta corrente**.

## Baseline conservata

- Aggiorna tutto: importazione, validazione, deduplicazione, scrittura dati, analitiche, dashboard e report; salvataggio finale.
- Stati protocollo presenti ma non mappati: WARNING registrato, protocollo conservato e aggiornamento non bloccato per questo motivo.
- Correzioni PDF/XLS, accodamento staging, confronto dei testi simili a date e anteprima PA3.
- Debug del salvataggio mantenuto, compreso DisplayAlerts al salvataggio finale.
- Esportazione ordinaria come bozza tecnica; nessuna autorizzazione Gold implicita.
- HOME 1.1: collegamenti interni ai fogli, senza dipendenza dal nome del workbook.

## Documenti correnti

- [Congelamento e perimetro Alpha](releases/alpha-1.0.3-home1.1/FREEZE.md)
- [Identità dell'archivio Alpha](releases/alpha-1.0.3-home1.1/artifact.json)
- [Hash dei componenti Alpha](releases/alpha-1.0.3-home1.1/COMPONENTS.sha256)
- [Release corrente M10.4/M11.4](releases/beta-m10.4-m11.4/README.md)
- [QA sistemico M10.4/M11.4](releases/beta-m10.4-m11.4/QA_SUMMARY.md)
- [Matrice di integrazione](releases/beta-m10.4-m11.4/SYSTEM_INTEGRATION_STATUS.md)
- [Piano beta B0](docs/BETA_B0.md)
- [Verifica dell'archivio Alpha ricevuto](tools/verify_frozen_archive.py)

Non reinstallare o reimportare moduli nella copia Alpha funzionante. Conservare l'Alpha separatamente e svolgere i test beta su copie prive di dati da pubblicare.
