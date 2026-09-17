# WebArch Beta — M10.5 / M11.5 — STEP 10A

## Stato

**OFFLINE FUNCTIONAL AUDIT PASS / RC BLOCKED — gate Microsoft Excel/VBE e workbook reale ancora PENDING.**

Questo checkpoint è la continuazione cumulativa del record GitHub STEP 6B. Non riscrive né sostituisce la Alpha congelata `EASYONE MINIMAL 1.0.3 + HOME 1.1` e non promuove la Beta a RC.

## Catena di remediation dopo STEP 6B

### STEP 7A — PDF hardening

Chiusi offline:

- `PDF-P2-01`: filename PDF non più hardcoded a M10.3; deriva da `OP_M10_Version()`;
- `PDF-P2-02`: nuovo `IsCurrent` subito dopo `GetSaveAsFilename` e subito prima di `ExportAsFixedFormat`.

Esito: **18/18 suite PASS**; verifier PDF **18/18 PASS**.

### STEP 8A — macro surface classification

Classificate **135/135** `Public Sub` runtime senza argomenti obbligatori, con **zero modifiche VBA**:

- 56 `SUPPORTED_UI_BOUND`;
- 25 `BETA_ADMIN_SETUP_STATUS`;
- 18 `NATIVE_TEST_GATE`;
- 14 `LEGACY_UNBOUND`;
- 10 `RECOVERY_ADMIN`;
- 6 `SUPPORT_NAV_UNBOUND`;
- 5 `SUPPORT_DIAGNOSTIC`;
- 1 `INTERNAL_CROSS_MODULE_API`.

La riduzione della superficie non viene eseguita offline: una `.xlsm` reale può contenere binding `Shape.OnAction` persistenti non visibili nel solo sorgente VBA.

### STEP 8B — real workbook binding gate

Audit read-only preparato; template Alpha e generatori VBA non mostrano binding verso i 20 candidati, ma la `.xlsm` operativa non è disponibile nel corpus verificato.

Stato `SURFACE-P1-01`: **CLASSIFIED_OFFLINE / REAL_WORKBOOK_BINDING_GATE_PENDING**.

Nessuna macro è stata resa `Private`, rimossa o nascosta.

### STEP 9 — total offline re-audit

Ricostruzione da zero della Beta cumulativa e audit indipendente:

- **18/18 suite PASS**;
- **320/320 controlli audit PASS**;
- 91/91 moduli conformi al manifest;
- 0 nomi `VB_Name` duplicati;
- 0 procedure pubbliche duplicate.

Confermati chiusi offline: P0 writer/recovery/bootstrap, snapshot locale, verifier procedure-scoped, M11 fail-safe, button binding, EASYONE/reentrancy e PDF hardening.

### STEP 10A — versioning e nomenclatura

Normalizzati senza cambiare firme API o flussi dati:

- runtime M10 → `BETA-M10.5-REMEDIATION-1.0.0`;
- M4H2 → `BETA-M4H2-SAFETY-GUARD-1.1.0`;
- messaggi/gate M10/M11 coerenti con M10.5/M11.5;
- riferimenti correnti M4H3/M4H4 eliminati, mantenendo invariata l'API pubblica `OP_M4H2_*`.

Esito:

- suite adattata: **19/19 PASS**;
- gate STEP10A: **24/24 PASS**;
- static audit: **643/643 PASS**;
- 91 moduli, 7 modificati;
- macro surface invariata;
- cold overlay: **91/91 byte-identici** e **19/19 suite PASS**.

## Moduli modificati in STEP 10A

- `100_modOP_M11_4_NativeAcceptance.bas` — SHA-256 `dd4f8c7a9fdfea9a9b660550221fc28cb534c7edc3e4becdcd48a7f391cdc907`
- `81_modOP_M10_Core.bas` — SHA-256 `9f9c335653b4adc639503748c07a3d8c3087b8691aed0fc4a76f4a3a28f4f81e`
- `82_modOP_M10_Update.bas` — SHA-256 `826242da2212b5d488ff4db097fe2dd5e9b355323e419eb573b26294145c31b0`
- `83_modOP_M10_Report.bas` — SHA-256 `f3923fef224ff47bbdd1ddcda2b405c74be65397fb5f5395b44a189732a2bbc7`
- `84_modOP_M10_Main.bas` — SHA-256 `cdb4965ae0886505b87c541f955696297cb4cdf0ea526ddcce8377774ae9914a`
- `85_modOP_M4H2_Guard.bas` — SHA-256 `27c827abf3107841eefe6ec4abd53014dab273efa1b3959c4321b42305b53b46`
- `99_modOP_M10_NativeTests.bas` — SHA-256 `3d3442f5f9f7feb5e4b6de76f6e3b86cbeec1fb87a2ebb6491e3c9d678390c88`

## Checksum package Library

`WebArch_M10_5_STEP10A_VERSIONING_NAMING.zip`

SHA-256: `31d68fb7eb6388f5d44c12ce5b79a51f852a9823323ad88671a5e26f58dc03fe`

Il package ZIP è archiviato nella Project Library. Il repository pubblico registra checksum, rapporti e diff; non viene dichiarato che il package binario o la `.xlsm` operativa siano scaricabili da GitHub.

## Finding correnti

### Chiusi offline

- `SYS-P0-01`, `SYS-P0-02`, `SYS-P0-03`;
- `SYS-P1-01`;
- `QA-P1-01`;
- `TEST-P1-01`;
- `UI-P1-01`, `UI-P1-02`;
- `PDF-P2-01`, `PDF-P2-02`;
- `DOC-P2-M10-RUNTIME-VERSION`;
- `DOC-P2-M4-NAMING`.

### Ancora aperti / bloccanti

- `SURFACE-P1-01`: **PENDING_REAL_WORKBOOK_GATE** — classificazione completa, riduzione vietata finché non viene auditata la `.xlsm` reale;
- `SYS-P2-CONCURRENCY-SCOPE`: **OPEN_OPTIMIZATION** — token globale PROTOCOLLO sicuro ma conservativo; non è un lost-update;
- `NATIVE-GATE`: **PENDING** — Excel/VBE reale;
- `RC-PACKAGE`: **PENDING** — manca un singolo package RC cumulativo installabile e ricertificato.

## Gate nativi obbligatori prima della RC

1. `Debug > Compila VBAProject`;
2. install/upgrade reale su copia Beta;
3. M3↔M4 stale/reload/apply reali;
4. Workbook events / `cMinimalSaveEvents`;
5. Save → close → reopen;
6. FSC identico / nuovo / bloccante;
7. PDF `ExportAsFixedFormat` reale;
8. due workbook aperti e dispatch `Shape.OnAction`;
9. performance/memoria su dataset reale;
10. audit read-only STEP8B sulla vera `.xlsm` prima di ridurre la macro surface.

## File di evidenza pubblicati in questo record

- `STEP7A_P2_PDF_HARDENING_REPORT.md`
- `VBA_DIFF_STEP7A_PDF.patch`
- `STEP8A_MACRO_SURFACE_REPORT.md`
- `STEP8B_WORKBOOK_BINDING_AUDIT_REPORT.md`
- `STEP9_TOTAL_OFFLINE_REAUDIT_REPORT.md`
- `STEP9_FINDINGS.json`
- `STEP10A_VERSIONING_NAMING_REPORT.md`
- `VBA_DIFF_STEP10A.patch`
- `STEP10A_STATIC_SUMMARY.json`
- `PACKAGE_SHA256.txt`

La PR `beta -> main` deve restare **Draft / DO NOT MERGE** fino alla chiusura del gate RC.