# WebArch Beta — M10.5 / M11.5 — STEP 6B

## Stato

**Remediation checkpoint offline verificato — NON RC congelata. Gate Microsoft Excel/VBE ancora PENDING.**

Questo record sostituisce come stato corrente la precedente milestone M10.4/M11.4. La milestone M10.4/M11.4 resta conservata come storico, ma il successivo audit totale di flusso ha evidenziato difetti sistemici poi corretti negli STEP 1–6B.

## Correzioni cumulative chiuse offline

- P0-01: protezione lost-update M3/M4 nel vero writer M4.
- P0-02: RICARICA M4 recuperabile con baseline stantia/assente.
- P0-03: bootstrap/upgrade senza dipendenza prematura dalle baseline.
- P1-01: snapshot locale M3/M4 separato dalla baseline esterna M2.
- P1-03: M11 fail-safe dopo mutazioni M2 e test nativo della distinzione edit locale/cambiamento esterno.
- P1-04: pulsanti M10 workbook-qualified.
- P1-05: hardening EASYONE/reentrancy per `OP_M10_INSTALLA`, `OP_M10_APRI_SINTESI`, `OP_M10_ESPORTA_PDF` e `OP_M4_AGGIUNGI_CAMPO`.

## STEP 6B — file VBA modificati rispetto a STEP 6A

Solo tre moduli:

- `83_modOP_M10_Report.bas` — SHA-256 `7e923284942edd6f759c8c317f39e457fcd35696c08af7b0009a3ef3fdd516bf`
- `84_modOP_M10_Main.bas` — SHA-256 `9eb4d9f0e7b406c72e615a188df67eacb6a26338066d4cf69e05d64e6bb8409a`
- `60_modOP_M4_Main.bas` — SHA-256 `bf98e6c462528c802c75e2fc68feb3e7cc11d132a70efa72f592b41c570e0d17`

Il repository conserva per questo checkpoint il diff VBA esatto e gli hash dei file modificati. La consegna completa verificata resta archiviata nella Project Library WebArch; nessun workbook operativo o dato reale viene pubblicato nel repository pubblico.

## QA offline corrente

- suite cumulative: **17/17 PASS**;
- release verifier: **92/92 PASS**;
- P0-01: **10/10 PASS**;
- P0-02: **22/22 PASS**;
- P0-03: **52/52 PASS**;
- P1-01 snapshot locale: **48/48 PASS**;
- P1-03 M11 fail-safe: **37/37 PASS**;
- P1-04 workbook-qualified buttons: **14/14 PASS**;
- P1-05 reentrancy: **40/40 PASS**;
- compile-oriented/static sanity STEP 6B: **559/559 PASS**;
- M10 models: **59/59 PASS**;
- functional-flow models: **31/31 PASS**;
- system-integration models: **13/13 PASS**;
- random integration: **5.000 iterazioni / 20.000 controlli PASS**;
- system-integration mutations: **14/14 PASS**;
- shared-state audit: **15/15 PASS**.

Cold overlay: STEP 6A pulito + STEP 6B → **17/17 suite PASS** e **15/15 moduli VBA byte-per-byte identici** al workspace auditato.

## Checksum consegna STEP 6B

`WebArch_M10_5_STEP6B_P1_05_REENTRANCY.zip`

SHA-256: `021a55369e87c56419bbec22f29658052552e990704ba3bdad4efedb7b157e0f`

## Gate nativo ancora obbligatorio

Prima di qualunque freeze RC:

1. `Debug > Compila VBAProject` in Microsoft Excel;
2. test reale EASYONE/reentrancy durante `AGGIORNA_TUTTO` attraversando i `DoEvents`;
3. prova con due copie WebArch aperte per il dispatch `Shape.OnAction`;
4. Save / close / reopen;
5. PDF `ExportAsFixedFormat` e controllo validità fino all'export;
6. eventi Workbook/cMinimalSaveEvents;
7. prove runtime dei ListObject e performance su dataset reale.

## File di questo record

- `STEP6B_P1_05_REPORT.md` — rapporto tecnico.
- `VBA_DIFF_STEP6B.patch` — diff esatto dei tre moduli modificati.
- `QA_SUMMARY.md` — riepilogo QA.
- `RUN_ALL_QA_RESULT.json` — risultato strutturato della suite cumulativa.
- `SHA256SUMS.txt` — hash dei file della consegna STEP 6B.
- `PACKAGE_SHA256.txt` — checksum dell'archivio consegnato.

L'Alpha EASYONE MINIMAL 1.0.3 + HOME 1.1 resta congelata e non viene modificata da questa release.