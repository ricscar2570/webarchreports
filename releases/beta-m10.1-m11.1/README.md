# WebArch — M10.1 Audit Remediation + M11.1 Native Acceptance

Data: 15 settembre 2026

## Cosa contiene

### M10_1
- `81_modOP_M10_Core.bas`
- `82_modOP_M10_Update.bas`
- `83_modOP_M10_Report.bas`
- `84_modOP_M10_Main.bas`
- `99_modOP_M10_NativeTests.bas`
- rapporto, gate Excel, matrice remediation, QA statico e model tests.

### M11_1
- `100_modOP_M11_1_NativeAcceptance.bas`
- evidenze e gate nativo.

### M4H1
- `M4H1_HARDENING_REQUIRED.txt`

## Stato

- QA statico M10.1/M11.1: **179/179 PASS**
- Model tests: **26/26 PASS**
- VBE compile: **NON ESEGUITO**
- Runtime Excel: **NON ESEGUITO**
- Save/reopen: **NON ESEGUITO**
- Performance reale: **NON ESEGUITA**
- RC1: **NON CONGELATA**

## Regola di installazione

Questi moduli sostituiscono soltanto il vecchio M10/M11. Alpha, Fascicoli B1 e M1–M9 non devono essere reinstallati o modificati da questo pacchetto.

Se il workbook contiene già il vecchio M10:
1. fai una copia del workbook;
2. rimuovi dal VBE soltanto `modOP_M10_Core`, `modOP_M10_Update`, `modOP_M10_Report`, `modOP_M10_Main`;
3. importa i quattro file M10.1;
4. `Debug > Compila VBAProject`;
5. solo dopo esegui `OP_M10_INSTALLA`.

Il test `99_modOP_M10_NativeTests.bas` va importato solo in una copia sacrificabile `WebArch_TEST_OP_M10...` e va chiusa senza salvare.

Il test `100_modOP_M11_1_NativeAcceptance.bas` va importato solo in una copia `WebArch_TEST_OP_M11...`.

## Nota M4H1

Il micro-hardening M4 resta requisito prima di RC1. Non è auto-applicato perché il sorgente M4 byte-esatto non era disponibile nel runtime di packaging. Non usare macro che auto-modificano il VBProject.

## Principio di sicurezza

M10.1 è intenzionalmente fail-closed: API di preflight non disponibile, fingerprint non valido, `MINIMAL_DiscardRequired`, refresh core fallito o fallimento dopo scrittura Alpha/FSC impediscono la certificazione corrente e il Save finale quando previsto.
