# WebArch — STEP 7A P2 PDF Hardening

Data: 16 settembre 2026
Base: M10.5/M11.5 remediation cumulativa STEP 6B
Stato: OFFLINE VERIFIED — gate Microsoft Excel ancora PENDING

## Obiettivi

Chiudere esclusivamente i finding PDF:

- PDF-P2-01: filename rimasto a `M10_3`.
- PDF-P2-02: stato non ricontrollato dopo il dialogo Save As e immediatamente prima di `ExportAsFixedFormat`.

## Modifica VBA

È stato modificato un solo modulo:

`M10_4/83_modOP_M10_Report.bas`

### Provenienza filename

Il filename non contiene più un tag hardcoded `M10_3`.
La versione viene derivata da `OP_M10_Version()` e normalizzata per il filesystem:

`WebArch_REPORT_OPERATIVO_<runtime-version>_yyyymmdd_hhnnss.pdf`

Con il runtime corrente, un nome tipico è:

`WebArch_REPORT_OPERATIVO_BETA_M10_4_SYSTEM_INTEGRATION_HARDENING_1_0_0_20260916_....pdf`

Questa scelta evita di introdurre in STEP 7A un bump del contratto runtime: la remediation è etichettata M10.5, mentre `OP_M10_VERSION` resta formalmente M10.4 fino allo step di versioning/migrazione dedicato. Il filename seguirà automaticamente qualunque futuro bump di `OP_M10_Version()`.

### Validità temporale dell'export

La sequenza è ora:

1. discard gate;
2. pending-edits gate;
3. `IsCurrent` iniziale;
4. ricostruzione output;
5. `IsCurrent` dopo ricostruzione;
6. `GetSaveAsFilename`;
7. **nuovo `IsCurrent` immediatamente dopo il ritorno dal dialogo**;
8. gestione estensione/overwrite;
9. discard gate;
10. `OP_REPORT.Calculate`;
11. **nuovo `IsCurrent` immediatamente prima di `ExportAsFixedFormat`**;
12. export PDF.

Questo chiude sia la finestra temporale del Save As sia quella successiva dovuta a overwrite prompt/calcolo.

## Regressioni

Baseline pre-patch: **17/17 suite PASS**.

Dopo STEP 7A:

- release verifier: **97/97 PASS**;
- PDF hardening verifier: **18/18 PASS**;
- P0-01: 10/10;
- P0-02: 22/22;
- P0-03: 52/52;
- P1-01 snapshot: 48/48;
- P1-03 M11 fail-safe: 37/37;
- P1-04 button binding: 14/14;
- P1-05 reentrancy: 40/40;
- STEP6B static sanity: 559/559;
- M10 models: 59/59;
- functional flow: 31/31;
- system integration: 13/13;
- random integration: 20.000/20.000;
- integration mutations: 14/14;
- shared-state: 15/15;
- **suite totali: 18/18 PASS**.

Mutation test PDF: rimozione dei gate, reintroduzione `M10_3` e rimozione della derivazione da runtime version vengono rilevate.

## Diff source

15 moduli VBA confrontati con STEP 6B:

- **14 identici**;
- **1 modificato:** `83_modOP_M10_Report.bas`.

Hash STEP 6B del modulo:
`7e923284942edd6f759c8c317f39e457fcd35696c08af7b0009a3ef3fdd516bf`

Hash STEP 7A:
`6849f076e1a43befac57bf577f84bc41d918be481ba2bbc4e80e34dbb73cbcb3`

## Cold overlay

Una copia pulita STEP 6B ha ricevuto esclusivamente il patch STEP 7A.

Risultato:

- **18/18 suite PASS**;
- **15/15 moduli VBA byte-identici** al workspace auditato.

## Limiti

Non è stato eseguito Microsoft Excel reale. Restano obbligatori:

- `Debug > Compila VBAProject`;
- `GetSaveAsFilename` reale;
- variazione dello stato mentre il dialogo è aperto;
- `ExportAsFixedFormat` reale;
- controllo presenza/dimensione PDF su filesystem.

## Stato finding

- PDF-P2-01: **CLOSED_OFFLINE_STEP7A**.
- PDF-P2-02: **CLOSED_OFFLINE_STEP7A**.
- Native PDF gate: **PENDING**.
