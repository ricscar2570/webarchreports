# WebArch — STEP 8A Macro Surface Classification
## 16 settembre 2026

**Stato: CLASSIFICATION COMPLETE / ZERO VBA CHANGES.**

Obiettivo: classificare l’intera superficie macro runtime prima di rendere Private, nascondere o rimuovere qualsiasi entry point. Questo step è intenzionalmente read-only per evitare regressioni.

## Perimetro verificato

- 91 file VBA distribuiti ricostruiti complessivamente (Alpha + B1 + M1–M11 + remediation STEP 1–7A).
- 4 moduli di setup/migrazione distribuiti ma non importati nel workbook runtime sono esclusi dalla macro surface operativa.
- 87 moduli runtime effettivi nel perimetro.
- **135 Public Sub senza argomenti obbligatori** in moduli standard non `Option Private Module`: il conteggio coincide con l’audit applicativo precedente.
- 12 macro aggiuntive appartengono ai soli moduli setup/migrazione esclusi dal runtime; includerle produrrebbe 147 e sarebbe un perimetro errato.

## Classificazione 135/135

| Categoria | Conteggio | Politica STEP 8A |
|---|---:|---|
| `SUPPORTED_UI_BOUND` | 56 | **KEEP PUBLIC** — legate a hotspot/shape dell’interfaccia supportata. |
| `BETA_ADMIN_SETUP_STATUS` | 25 | **KEEP PUBLIC UNTIL NATIVE GATE** — installazione, apertura, stato e supporto Beta. |
| `NATIVE_TEST_GATE` | 18 | **KEEP PUBLIC UNTIL NATIVE GATE** — necessarie per il collaudo Excel/VBE. |
| `LEGACY_UNBOUND` | 14 | **CANDIDATE RETIRE/HIDE**, ma solo dopo inventario `OnAction` del workbook reale. |
| `RECOVERY_ADMIN` | 10 | **KEEP PUBLIC PROTECTED** — recovery/performance/emergenza. |
| `SUPPORT_NAV_UNBOUND` | 6 | **KEEP PENDING WORKBOOK INVENTORY** — navigazioni secondarie non legate dal setup corrente. |
| `SUPPORT_DIAGNOSTIC` | 5 | **KEEP PUBLIC PROTECTED** — diagnostica/supporto e routing cross-module. |
| `INTERNAL_CROSS_MODULE_API` | 1 | **KEEP PUBLIC** — `MINIMAL_RequireDiscard`, API interna usata dal runtime. |

## Perché non ho ancora reso Private alcuna macro

Il codice sorgente non è l’unica fonte di binding: un workbook `.xlsm` può contenere `Shape.OnAction` persistenti che non compaiono come riferimenti testuali nel VBA corrente. Una macro con zero riferimenti nel sorgente può quindi essere ancora richiamata da una forma presente nel workbook operativo. Privatizzarla offline senza inventario nativo potrebbe rompere un pulsante apparentemente “senza dipendenze”.

Inoltre `Option Private Module` non va applicato indiscriminatamente ai moduli misti: diversi moduli contengono insieme macro UI supportate e backend/admin. Nascondere l’intero modulo rischierebbe di impedire il dispatch delle Shape.

## Candidati legacy reali — 14

- `G15OpenHardeningView` — `23_modHardening.bas`
- `G15OpenPrivacyMatrix` — `23_modHardening.bas`
- `G15RunHardeningCheck` — `23_modHardening.bas`
- `EseguiCollaudo` — `25_modGoldValidation.bas`
- `G13OpenGoldView` — `25_modGoldValidation.bas`
- `G13OpenEvidenceView` — `25_modGoldValidation.bas`
- `G13BindUatToCurrentDataset` — `25_modGoldValidation.bas`
- `G13ExportReport` — `25_modGoldValidation.bas`
- `ApriEvidenzeGold` — `27_modNavigation.bas`
- `ApriCollaudoGold` — `27_modNavigation.bas`
- `ApriHardening` — `27_modNavigation.bas`
- `ApriPrivacyMatrix` — `27_modNavigation.bas`
- `ApriPerformanceRecovery` — `27_modNavigation.bas`
- `ApriRecoveryG16` — `27_modNavigation.bas`

Questi comandi appartengono soprattutto alle vecchie superfici Gold/G15/G16 e alle relative navigazioni. Il setup MINIMAL corrente non li lega a hotspot supportati; **non vengono però modificati nello STEP 8A**.

## Navigazioni secondarie non legate — 6

- `ApriIstruzioni` — `27_modNavigation.bas`
- `ApriFontiUfficiali` — `27_modNavigation.bas`
- `ApriDiagnosticaFonti` — `27_modNavigation.bas`
- `ApriStorico` — `27_modNavigation.bas`
- `ApriGiacenze` — `27_modNavigation.bas`
- `TornaHome` — `27_modNavigation.bas`

Anche queste restano intatte fino all’inventario reale delle Shape/OnAction.

## Test nativi — 18

Queste macro non devono essere nascoste prima del gate Excel perché servono proprio a eseguire il collaudo. Dopo la certificazione nativa sarà possibile valutare di spostare i moduli test fuori dal workbook di produzione o renderli non esposti.

## Strategia di riduzione sicura proposta per STEP 8B

1. Eseguire su **una copia Beta reale** un inventario read-only di tutte le `Shape.OnAction`, hyperlink, controlli modulo e altri binding macro.
2. Confrontare i nomi trovati con le 135 macro classificate.
3. Marcare `LEGACY_UNBOUND` e `SUPPORT_NAV_UNBOUND` come realmente inutilizzate solo se assenti anche dal workbook reale.
4. Durante il native gate mantenere pubbliche tutte le `NATIVE_TEST_GATE`.
5. Solo dopo, effettuare una riduzione in piccoli gruppi con compile VBE + regressione completa dopo ogni gruppo. Nessuna conversione massiva.

## QA e integrità

- Macro-surface verifier: **17/17 PASS**.
- Classificazione: **135/135**, 0 non classificati, 0 nomi duplicati.
- UI supported: **56/56 ancora pubbliche**.
- Source integrity: **91/91 file VBA identici**, 0 modificati durante STEP 8A.
- QA cumulativa STEP 7A rieseguita dopo la classificazione: **18/18 suite PASS**.
- Aggregate source fingerprint: `72baec1d921e330103462ae428c2abd5ea687dc87d57c275b4e9bfbdd0a9a079`.

## Decisione

**SURFACE-P1-01 non è ancora chiuso come riduzione**, ma il suo inventario/classificazione è ora completo e verificato. Ridurre la superficie prima dell’inventario nativo del workbook sarebbe più rischioso del beneficio e violerebbe l’obiettivo di evitare regressioni.
