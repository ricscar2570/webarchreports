# Beta STEP10A R4.8 — stato validato

Data: 21 settembre 2026.

## Risultato corrente

La R4.8 è la prima versione STEP10A di questo ciclo che ha superato la compilazione VBA nativa e il percorso di configurazione appena collaudato senza l'errore bloccante precedente. La conferma proviene dal test dell'utente su Microsoft Excel Desktop.

Questo stato significa **baseline beta funzionante per il percorso verificato**, non assenza assoluta di bug in ogni scenario.

## Sequenza dei difetti chiusi

1. **R4.4 — template OOXML**: corrette 9 definizioni di tabella con 17 nomi di colonne non coerenti con le intestazioni reali. Il template torna ad aprirsi nativamente in Excel mantenendo 88 fogli e 79 tabelle.
2. **R4.5/R4.6 — namespace VBA**: eliminate collisioni case-insensitive fra costanti e funzioni, inclusa `FSC_SHEET / FSC_Sheet`. Compilazione VBA nativa successivamente PASS.
3. **R4.7 — M8 runtime**: sostituito l'uso bloccante di `Shape.PrintObject` nel percorso OP_INCIDENTI. M8 passa le postcondizioni native.
4. **R4.8 — OnAction**: il validatore ora confronta semanticamente workbook e macro, accettando le forme Excel equivalenti quoted/unquoted e continuando a rifiutare macro non qualificate, workbook diversi o macro fuori whitelist.

## Evidenza runtime

Nel collaudo nativo precedente a R4.8 risultavano PASS:
- installazione manuale completa;
- FSC;
- M2, M3, M4, M5, M6, M7;
- M8 dopo il fix R4.7;
- M9;
- M10.

Il solo blocco residuo era il falso negativo finale del validatore OnAction. Dopo la correzione R4.8 l'utente ha confermato che la versione non dà errore.

## Artefatti

Pacchetto completo:
`WebArch_STEP10A_R4_8_ONACTION_FIX_COMPLETO.zip`
SHA-256: `e62101ee45de12b8f7efdabeb53c69e5b4bccfeeff4f3315d7aaf9f2c3de1798`

Patch R4.7 → R4.8:
`WebArch_STEP10A_R4_8_ONACTION_FIX_PATCH_ONLY.zip`
SHA-256: `7d6bd39d610fab5e77af67e6d32c18064e45eab6b92425bb8171d30b16261fab`

I pacchetti completi restano consegnati separatamente; questo repository registra identità, stato e verifiche.
