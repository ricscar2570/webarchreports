# WebArch Reports — WebArch v1.1

WebArch è un'applicazione amministrativa Excel per il ciclo mensile di acquisizione, normalizzazione, controllo, storicizzazione e rendicontazione di protocolli e fascicoli.

## Stato corrente

**v1.1.1 — Builder hotfix / collaudo Windows in corso**

Il primo test reale su Excel/Windows del 1 settembre 2026 ha individuato un difetto nel builder v1.1.0: il template `.xlsx` poteva esporre `VBProject` ma restituire `VBComponents = null`. In PowerShell `$null.Count` viene valutato come `0`, producendo il log `Componenti iniziali: 0`; il successivo `VBComponents.Import()` falliva con `Impossibile chiamare un metodo su un'espressione con valore null`.

Il builder **v1.1.1** corregge il problema convertendo prima il template in `.xlsm`, chiudendolo e riaprendolo, quindi verificando esplicitamente `VBProject` e `VBComponents` prima dell'importazione dei moduli. Vedi `PATCH_NOTES_v1.1.1.md`.

## Funzioni previste

Il workbook operativo espone quattro comandi:

- **AGGIORNA TUTTO** — importazione universale, staging, deduplicazione, storico, KPI, criticità e aggiornamento report;
- **DIAGNOSTICA FONTI** — riconoscimento di file e fogli in base alle intestazioni, indipendentemente da nomi e ordine delle colonne;
- **ESEGUI COLLAUDO** — controlli strutturali, riconciliazioni, privacy e verbale associato all'ultimo aggiornamento;
- **ESPORTA REPORT** — produzione controllata di PDF, KPI, criticità e manifest, consentita soltanto dopo collaudo valido.

## Installazione / build

La build finale richiede Microsoft Excel desktop per Windows.

1. Estrarre il pacchetto WebArch completo in una cartella locale scrivibile.
2. Abilitare temporaneamente in Excel **Considera attendibile l'accesso al modello a oggetti dei progetti VBA**.
3. Chiudere Excel.
4. Assicurarsi di avere la versione corrente di `BUILD/Build-WebArch.ps1` (builder 1.1.1).
5. Eseguire `BUILD_WEBARCH.cmd`.
6. Il risultato atteso è `Cruscotto_Gabinetto_WebArch_v1.1.xlsm` e nel log deve comparire `Smoke test VBA: OK`.
7. Disabilitare nuovamente l'accesso programmatico al progetto VBA dopo la build.

## Stato di affidabilità

La v1.1.1 **non è ancora classificata RC2/produzione**. Prima dell'uso istituzionale devono essere completati: compilazione VBA reale, smoke test, prova di idempotenza, rollback controllato, collaudo su cartella di rete, verifica privacy degli output e due cicli mensili consecutivi.

La repository viene riallineata progressivamente al pacchetto sorgente completo; non usare la sola presenza dei file su GitHub come prova di collaudo runtime.
