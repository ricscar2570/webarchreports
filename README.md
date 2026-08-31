# WebArch Reports — WebArch v1.1

WebArch è un'applicazione amministrativa Excel per il ciclo mensile di acquisizione, normalizzazione, controllo, storicizzazione e rendicontazione di protocolli e fascicoli.

La release corrente è **v1.1.0 — Source release**. Il repository contiene il template Excel, i moduli VBA, il builder Windows, gli script di verifica, le fonti dimostrative, la documentazione tecnica e il pacchetto distributivo completo.

## Funzioni principali

Il file operativo espone quattro comandi:

- **AGGIORNA TUTTO** — importazione universale, staging, deduplicazione, storico, KPI, criticità e aggiornamento report;
- **DIAGNOSTICA FONTI** — riconoscimento di file e fogli in base alle intestazioni, indipendentemente da nomi e ordine delle colonne;
- **ESEGUI COLLAUDO** — controlli strutturali, riconciliazioni, privacy e verbale associato all'ultimo aggiornamento;
- **ESPORTA REPORT** — produzione controllata di PDF, KPI, criticità e manifest, consentita soltanto dopo collaudo valido.

## Creazione del workbook operativo

Il repository distribuisce un template `.xlsx` senza macro. Per generare il workbook operativo:

1. clonare o scaricare il repository in una cartella locale scrivibile;
2. verificare la presenza di Microsoft Excel desktop per Windows;
3. abilitare temporaneamente in Excel l'accesso attendibile al modello a oggetti dei progetti VBA;
4. chiudere Excel;
5. eseguire `BUILD_WEBARCH.cmd`;
6. eseguire `VERIFICA_BUILD.cmd`;
7. aprire `Cruscotto_Gabinetto_WebArch_v1.1.xlsm` e usare prima **DIAGNOSTICA FONTI** sulla cartella `ESEMPI_FONTI`.

Il builder non modifica le impostazioni di sicurezza di Excel e non scrive nel registro di Windows.

## Struttura

```text
BUILD/                  builder e verifica PowerShell
DOCUMENTAZIONE/         manuali, specifica, collaudo e migrazione
ESEMPI_FONTI/           dataset dimostrativi
VBA/                    dodici moduli e codice ThisWorkbook
Cruscotto_...xlsx       template completo senza macro
dist/                   pacchetto ZIP distributivo e checksum
```

## Stato del rilascio

La release comprende validazione statica e test predisposti, ma prima dell'uso istituzionale richiede:

- compilazione e prova reale in Excel Windows;
- collaudo su cartella locale e di rete;
- prova con operatore non tecnico;
- verifica del rollback;
- due cicli mensili consecutivi;
- firma digitale del progetto VBA tramite certificato dell'Amministrazione;
- migrazione verificata dei dati RC1.

Consultare `LEGGIMI_PRIMA.txt`, `DOCUMENTAZIONE/MANUALE_OPERATIVO.md` e `VALIDAZIONE_STATICA.txt` prima dell'installazione.
