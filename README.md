# WebArch — EASYONE MINIMAL

## Stato del progetto

**Alpha congelata: EASYONE MINIMAL 1.0.3 + HOME 1.1.**

Il 14 settembre 2026 il responsabile del progetto ha confermato: «funziona tutto come previsto», richiedendo il congelamento dell'alpha e l'avvio della beta. La conferma riguarda il suo utilizzo in Excel; non viene trasformata in una certificazione automatica di tutti i casi di test.

La Beta parte dagli stessi componenti produttivi dell'Alpha e resta strettamente additiva: nessuna riscrittura dell'Alpha, nessuna rimozione del debug e nessun ritorno alla macchina transazionale. La versione interna VBA dell'Alpha resta EASYONE-MINIMAL-1.0.3; HOME resta 1.1.

## Beta corrente

**M10.5 / M11.5 — remediation cumulativa fino a STEP 6B — OFFLINE VERIFIED, RC non congelata.**

La precedente M10.4/M11.4 è conservata nel repository come milestone storica, ma il successivo audit totale di flusso l'ha dichiarata BLOCKED. Gli STEP 1–6B hanno corretto in sequenza i difetti rilevati senza modificare l'Alpha congelata.

Correzioni cumulative chiuse offline:

- protezione lost-update M3/M4 nel vero writer M4;
- RICARICA M4 recuperabile con baseline stantia/assente;
- bootstrap upgrade indipendente dalle nuove baseline ancora inesistenti;
- snapshot locale M3/M4 separato dalla baseline esterna M2;
- M11 fail-safe dopo mutazioni M2;
- pulsanti M10 qualificati al workbook proprietario;
- hardening EASYONE/reentrancy per `OP_M10_INSTALLA`, `OP_M10_APRI_SINTESI`, `OP_M10_ESPORTA_PDF` e `OP_M4_AGGIUNGI_CAMPO`.

Evidenza offline corrente:

- suite cumulative: **17/17 PASS**;
- release verifier: **92/92 PASS**;
- P0-01: **10/10 PASS**;
- P0-02: **22/22 PASS**;
- P0-03: **52/52 PASS**;
- P1-01 snapshot locale: **48/48 PASS**;
- P1-03 M11 fail-safe: **37/37 PASS**;
- P1-04 binding multi-workbook: **14/14 PASS**;
- P1-05 reentrancy: **40/40 PASS**;
- compile-oriented/static sanity STEP 6B: **559/559 PASS**;
- M10 models: **59/59 PASS**;
- functional-flow models: **31/31 PASS**;
- system-integration models: **13/13 PASS**;
- random integration: **5.000 iterazioni / 20.000 controlli PASS**;
- integration mutations: **14/14 PASS**;
- shared-state audit: **15/15 PASS**;
- cold overlay: **17/17 suite PASS**, **15/15 moduli VBA byte-per-byte identici** al workspace auditato.

Il pacchetto STEP 6B verificato ha SHA-256:

`021a55369e87c56419bbec22f29658052552e990704ba3bdad4efedb7b157e0f`

### Gate ancora aperto

Microsoft VBE/Excel reale non è ancora stato eseguito per questa remediation. Prima del freeze RC restano obbligatori: compilazione VBAProject, reentrancy reale durante `AGGIORNA_TUTTO`, prova con due copie WebArch aperte per `Shape.OnAction`, Workbook events, Save/close/reopen, PDF runtime e performance su dataset reale.

## Disponibilità dei file

Il repository pubblico contiene il registro del congelamento Alpha e i record delle milestone Beta. Per STEP 6B pubblica il rapporto tecnico, il diff VBA esatto, gli hash dei tre moduli modificati, l'evidenza QA e il checksum del pacchetto. La consegna completa verificata resta archiviata nella Project Library del progetto WebArch.

Il file `.xlsm` operativo con dati dell'utente non è pubblicato e non deve essere caricato su questo repository pubblico.

I file `BUILD/`, `BUILD_WEBARCH.cmd` e `PATCH_NOTES_v1.1.1.md` appartengono al builder storico pre-MINIMAL: sono conservati per tracciabilità, **non sono la procedura di installazione della beta corrente**.

## Baseline conservata

- Aggiorna tutto: importazione, validazione, deduplicazione, scrittura dati, analitiche, dashboard e report; salvataggio finale.
- Stati protocollo presenti ma non mappati: WARNING registrato, protocollo conservato e aggiornamento non bloccato per questo motivo.
- Correzioni PDF/XLS, accodamento staging, confronto dei testi simili a date e anteprima PA3.
- Debug del salvataggio mantenuto, compreso DisplayAlerts al salvataggio finale.
- Esportazione ordinaria come bozza tecnica; nessuna autorizzazione Gold implicita.
- HOME 1.1: collegamenti interni ai fogli, senza dipendenza dal nome del workbook.

## Documenti correnti

- [Release corrente M10.5/M11.5 STEP 6B](releases/beta-m10.5-step6b/README.md)
- [QA STEP 6B](releases/beta-m10.5-step6b/QA_SUMMARY.md)
- [Rapporto tecnico STEP 6B](releases/beta-m10.5-step6b/STEP6B_P1_05_REPORT.md)
- [Diff VBA STEP 6B](releases/beta-m10.5-step6b/VBA_DIFF_STEP6B.patch)
- [Risultato suite cumulativa](releases/beta-m10.5-step6b/RUN_ALL_QA_RESULT.json)
- [Checksum pacchetto STEP 6B](releases/beta-m10.5-step6b/PACKAGE_SHA256.txt)
- [Milestone storica M10.4/M11.4](releases/beta-m10.4-m11.4/README.md)
- [Congelamento e perimetro Alpha](releases/alpha-1.0.3-home1.1/FREEZE.md)
- [Identità dell'archivio Alpha](releases/alpha-1.0.3-home1.1/artifact.json)
- [Hash dei componenti Alpha](releases/alpha-1.0.3-home1.1/COMPONENTS.sha256)
- [Piano beta B0](docs/BETA_B0.md)
- [Verifica dell'archivio Alpha ricevuto](tools/verify_frozen_archive.py)

Non reinstallare o reimportare moduli nella copia Alpha funzionante. Conservare l'Alpha separatamente e svolgere i test Beta su copie prive di dati da pubblicare.
