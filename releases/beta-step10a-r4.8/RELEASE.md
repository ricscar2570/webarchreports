# WebArch STEP10A R4.8

**Stato:** beta validata sul percorso nativo appena collaudato.

Data: 2026-09-21.

## Esito

La R4.8 chiude il falso negativo del gate finale `Shape.OnAction`. Excel può restituire il binding come `Workbook.xlsm!Macro` senza apici singoli attorno al nome del workbook; R4.7 richiedeva invece la stringa raw quoted e segnalava errore pur puntando allo stesso workbook e alla stessa macro.

R4.8 normalizza la rappresentazione e valida semanticamente:
- workbook esatto corrente;
- macro esplicitamente qualificata;
- macro presente nella whitelist UI.

Restano rifiutati riferimenti non qualificati, workbook differenti e macro sconosciute.

## Validazione

- 91 moduli VBA.
- Compilazione VBA nativa: PASS.
- M2–M10 e relative postcondizioni nel percorso di configurazione: PASS nel collaudo utente.
- M8 runtime: PASS.
- R4.8: utente conferma assenza dell'errore precedente.

L'evidenza runtime è una conferma sul percorso testato e non sostituisce una campagna completa di regressione su tutti gli scenari applicativi.

## Artefatti

Completo: `WebArch_STEP10A_R4_8_ONACTION_FIX_COMPLETO.zip`  
SHA-256: `e62101ee45de12b8f7efdabeb53c69e5b4bccfeeff4f3315d7aaf9f2c3de1798`  
Dimensione: 898339 byte.

Patch: `WebArch_STEP10A_R4_8_ONACTION_FIX_PATCH_ONLY.zip`  
SHA-256: `7d6bd39d610fab5e77af67e6d32c18064e45eab6b92425bb8171d30b16261fab`  
Dimensione: 41219 byte.
