# WebArch — STEP 8B — Workbook Binding Audit Gate
## 16 settembre 2026

## Stato

**READ-ONLY / NO VBA CHANGES / REAL XLSM PENDING**

Questo step non modifica nessun modulo VBA. Serve a eliminare il rischio principale prima di ridurre la macro surface: una Shape o un controllo persistito nel workbook reale può richiamare una macro anche quando il nome non compare in nessun altro sorgente VBA.

## Disponibilità del workbook reale

Nel Project e nell'intera Library non è presente una copia `.xlsm` WebArch autonoma. L'archivio Alpha congelato contiene `WebArch_PA0_BASE_TEMPLATE.xlsx`, non il workbook operativo già lavorato in Excel.

Per questo motivo **nessuna delle 20 macro candidate è stata resa Private, rimossa o nascosta**.

## Evidenza sul template Alpha

Audit diretto del package OOXML `WebArch_PA0_BASE_TEMPLATE.xlsx`:

- fogli: **88**;
- drawing objects: **4**, tutti grafici;
- attributi/callback macro non vuoti: **0**;
- hyperlink OOXML persistiti: **0**;
- formule `HYPERLINK(...)`: **0**;
- binding verso una delle 20 macro candidate: **0**.

Gli attributi `macro=""` dei quattro grafici sono vuoti e non costituiscono binding.

## Evidenza dai generatori VBA correnti

Sono stati analizzati tutti i **91** moduli `.bas/.cls` della Beta cumulativa STEP 7A.

- righe relative a creazione/binding di pulsanti individuate: **55**;
- `Application.Run` / dynamic invocation rilevanti: **7**;
- candidati STEP 8A: **20**;
- candidati usati come target da builder/installer correnti: **0**;
- candidati richiamati da `Application.Run`: **0**;
- candidati con binding nel template Alpha: **0**.

`modManualButtons` conferma esplicitamente il perimetro MINIMAL:

- controlli G15: **non installati**;
- controlli G16/recovery: **non installati**;
- GOLD/UAT: nessun hotspot installato nel blocco corrente.

Questo riduce fortemente il rischio, ma non sostituisce l'inventario della copia `.xlsm` reale.

## Matrice dei 20 candidati

`CANDIDATE_DECISION_MATRIX.csv` contiene, per ogni macro:

- categoria STEP 8A;
- riferimenti lessicali;
- binding-generator hit;
- `Application.Run` hit;
- template binding hit;
- stato workbook reale;
- decisione corrente.

Risultato attuale: **20/20 = PENDING_REAL_XLSM_BINDING_AUDIT**.

Nessun candidato presenta oggi un binding noto offline.

## Auditor read-only

È stato creato:

`WebArch_AUDIT_XLSM_BINDINGS_READONLY.ps1`

Proprietà di sicurezza:

1. non usa Excel COM;
2. non lancia Excel;
3. non esegue macro;
4. non apre il workbook in modalità applicativa;
5. usa `System.IO.Compression.ZipFile.OpenRead`;
6. analizza drawing, controls, ctrlProps, VML, worksheets e customUI;
7. rileva `macro=`, `onAction`, callback Ribbon e `FmlaMacro`;
8. rileva hyperlink location;
9. calcola SHA-256/dimensione/mtime prima e dopo;
10. fallisce se il file sorgente cambia.

Il verifier dello strumento passa **40/40** controlli.

> Nota: l'ambiente di sviluppo corrente è Linux e non contiene PowerShell, quindi l'esecuzione nativa del `.ps1` avverrà sul PC Windows dell'utente. La struttura/sicurezza dello script e la logica OOXML sono state verificate offline, inclusa l'esecuzione equivalente sul template Alpha.

## Integrità del sorgente

Confronto con `FULL_SOURCE_MANIFEST.json` dello STEP 8A:

- moduli attesi: **91**;
- moduli presenti: **91**;
- differenze SHA-256: **0**.

Quindi STEP 8B introduce **zero modifiche VBA**.

## QA regressione

- STEP 8A macro-surface verifier: **17/17 PASS**;
- STEP 7A QA cumulativa: **18/18 suite PASS**;
- STEP 8B auditor verifier: **40/40 PASS**;
- full source: **91/91 byte-identico** al manifest STEP 8A.

## Decisione

`SURFACE-P1-01` **non viene ancora chiuso**.

Stato corretto:

**CLASSIFIED_OFFLINE / REAL_WORKBOOK_BINDING_GATE_PENDING**

La privatizzazione massiva resta vietata. Dopo l'audit della `.xlsm` reale:

- `BOUND_DO_NOT_HIDE` → mantenere Public;
- `REVIEW_RAW_PART_HIT` → analisi manuale prima di qualsiasi cambio;
- `NO_PACKAGE_BINDING_FOUND` → incrociare con riferimenti VBA e semantica, poi eventualmente ridurre a piccoli gruppi con QA completa.

## Prossimo gate

Eseguire il tool sulla copia `.xlsm` reale e riportare almeno:

- `WORKBOOK_BINDING_AUDIT_SUMMARY.json`;
- `CANDIDATE_BINDING_STATUS.csv`.

Solo dopo questo gate sarà sicuro iniziare STEP 8C con modifiche alla visibilità delle macro.
