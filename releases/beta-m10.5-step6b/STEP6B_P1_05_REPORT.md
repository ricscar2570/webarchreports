# WebArch M10.5 — STEP 6B
## P1-05 — EASYONE reentrancy hardening

Data: 16 settembre 2026
Base: checkpoint cumulativo STEP 6A.

## Obiettivo
Impedire che entry point Beta che mutano o ricostruiscono stato possano essere eseguiti mentre EASYONE e' gia' occupato, in particolare durante `AGGIORNA_TUTTO` e le finestre di re-entrancy generate dai `DoEvents` della pipeline Alpha.

Lo STEP 6B chiude il finding applicativo relativo a:
- `OP_M10_INSTALLA`;
- `OP_M10_APRI_SINTESI`;
- `OP_M10_ESPORTA_PDF`;
- `OP_M4_AGGIUNGI_CAMPO`.

## Strategia adottata
La correzione non introduce nuovi lock e non modifica l'Alpha congelata. Riutilizza esclusivamente il guard EASYONE gia' canonico.

### M10 INSTALLA
`OP_M10_INSTALLA` e' ora un wrapper pubblico che:
1. acquisisce EASYONE con owner `M10_INSTALLA`;
2. richiama il core privato `OP_M10_INSTALLA_CORE`;
3. rilascia EASYONE in `Cleanup` anche in presenza di errore del wrapper.

La logica di installazione/migrazione preesistente e' rimasta nel core privato senza alterazioni funzionali. Non e' stato aggiunto un secondo `MINIMAL_PrepareWorkbook` o un nuovo `EnterControlledExecution`, per non cambiare la semantica dell'installatore.

### M10 APRI SINTESI
`OP_M10_APRI_SINTESI` acquisisce EASYONE per tutta la durata dell'aggiornamento banner/apertura e rilascia il guard in cleanup. La logica precedente e' isolata nel core privato `OP_M10_APRI_SINTESI_CORE`.

### M10 ESPORTA PDF
`OP_M10_ESPORTA_PDF` acquisisce EASYONE prima di qualunque precondizione/ricalcolo/dialogo SaveAs e mantiene il guard fino al termine del core `OP_M10_ESPORTA_PDF_CORE`. Questo e' importante perche' l'export contiene un dialogo utente e non deve lasciare una finestra in cui un'altra operazione WebArch possa partire.

Non e' stato usato `EnterControlledExecution` sull'export: avrebbe modificato `Application.Interactive` e sarebbe incompatibile con il dialogo `GetSaveAsFilename`.

### M4 NUOVO CAMPO
`OP_M4_AGGIUNGI_CAMPO` ora usa lo stesso envelope gia' collaudato dagli entry point mutanti M5/M7/M8:

`EASYONE_TryBeginOperation` -> `MINIMAL_PrepareWorkbook` -> `CaptureExcelState` -> `EnterControlledExecution` -> core -> `RestoreExcelState` -> `EASYONE_EndOperation`.

L'owner dedicato e' `OP_M4_ADD`.

## Regole preservate
- nessuna modifica all'Alpha congelata;
- nessun nuovo sistema transazionale;
- nessun richiamo M3<->M4 aggiuntivo;
- nessuna modifica ai moduli di concorrenza/snapshot;
- nessuna modifica a M11 fail-safe;
- nessuna modifica ai core M2/M5/M7/M8/M9;
- binding workbook-qualified dello STEP 6A preservato.

## Diff VBA
Sono stati confrontati 15 moduli VBA cumulativi rispetto a STEP 6A.

- 12 moduli: byte-per-byte invariati;
- 3 moduli modificati:
  - `M10_4/83_modOP_M10_Report.bas`;
  - `M10_4/84_modOP_M10_Main.bas`;
  - `PATCH/60_modOP_M4_Main.bas`.

Statistiche diff:
- 83: +61 / -0 righe;
- 84: +31 / -0 righe;
- 60: +35 / -1 righe.

Le aggiunte M10 sono prevalentemente wrapper/cleanup; i core precedenti sono conservati come procedure private.

## QA offline
Baseline STEP 6A prima delle modifiche:
- 15/15 suite PASS.

Dopo STEP 6B:
- release verifier: 92/92 PASS;
- P0-01 writer guard: 10/10 PASS;
- P0-02 reload recovery: 22/22 PASS;
- P0-03 upgrade/bootstrap: 52/52 PASS;
- P1-01 local snapshot: 48/48 PASS;
- P1-03 M11 fail-safe: 37/37 PASS;
- P1-04 workbook-qualified buttons: 14/14 PASS;
- P1-05 reentrancy verifier: 40/40 PASS;
- STEP 6B compile-oriented/static sanity: 559/559 PASS;
- M11 static sanity: 15/15 PASS;
- STEP 4 static sanity: 52/52 PASS;
- M10 models: 59/59 PASS;
- functional-flow models: 31/31 PASS;
- system-integration models: 13/13 PASS;
- random integration: 5.000 iterazioni / 20.000 controlli PASS;
- system-integration mutations: 14/14 PASS;
- shared-state audit: 15/15 PASS;
- suite complessive: **17/17 PASS**.

Il verifier STEP 6B comprende 8 mutation negative specifiche: rimozione acquire/release M10, rimozione acquire summary, rimozione release export, rimozione prepare/restore/release M4 e spostamento della scrittura M4 prima del guard.

## Cold overlay
E' stata ricostruita una base pulita STEP 6A a partire dalle release precedenti; successivamente e' stato applicato esclusivamente il patch STEP 6B.

Risultato:
- 17/17 suite PASS;
- 15/15 moduli VBA byte-per-byte identici al workspace auditato.

## Gate nativo ancora necessario
L'offline QA non sostituisce Microsoft Excel/VBE. Prima del freeze RC e' ancora necessario:
1. `Debug > Compila VBAProject`;
2. avviare `AGGIORNA_TUTTO` su una copia Beta di test sufficientemente grande da attraversare i `DoEvents`;
3. durante l'operazione tentare i comandi M10/M4 interessati e verificare che EASYONE li rifiuti senza mutare fogli/configurazione;
4. al termine verificare che EASYONE torni `FREE`;
5. verificare successivamente che i quattro comandi funzionino normalmente quando EASYONE e' libero;
6. Save/close/reopen della copia di test.

## Stato
`APP-P1-05 / UI reentrancy` = **CLOSED_OFFLINE_STEP6B**.

La certificazione finale resta subordinata al gate nativo Excel.