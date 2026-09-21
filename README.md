# WebArch — EASYONE MINIMAL

## Stato del progetto

**Beta corrente: STEP10A R4.8 — OnAction normalization fix.**

Il 21 settembre 2026 la R4.8 è stata eseguita in Microsoft Excel Desktop dopo il ciclo di correzioni STEP10A. La compilazione VBA nativa è passata e l'utente ha confermato che questa versione non presenta l'errore che bloccava la configurazione. Il risultato è una validazione runtime dell'attuale percorso di build/configurazione; non va interpretato come certificazione automatica di ogni possibile caso d'uso.

La baseline storica resta **EASYONE MINIMAL 1.0.3 + HOME 1.1**, congelata il 14 settembre 2026. STEP10A estende quella baseline con i moduli e gli installer beta tracciati nelle release R3–R4.8.

## Evidenze R4.8

- 91 componenti VBA caricati.
- Compilazione nativa `Debug > Compila VBAProject`: PASS.
- Installazione/configurazione nativa M2–M10 con postcondizioni: PASS nel collaudo utente.
- M8 `OP_INCIDENTI`: PASS dopo il fix R4.7.
- Validazione `Shape.OnAction`: corretta in R4.8 normalizzando la sintassi quoted/unquoted del nome workbook senza allentare whitelist e controllo del workbook.
- L'utente ha confermato che la R4.8 non dà errore nel percorso appena collaudato.

Pacchetto completo consegnato separatamente nella conversazione di progetto:

`WebArch_STEP10A_R4_8_ONACTION_FIX_COMPLETO.zip`

SHA-256: `e62101ee45de12b8f7efdabeb53c69e5b4bccfeeff4f3315d7aaf9f2c3de1798`

Patch R4.7 → R4.8:

`WebArch_STEP10A_R4_8_ONACTION_FIX_PATCH_ONLY.zip`

SHA-256: `7d6bd39d610fab5e77af67e6d32c18064e45eab6b92425bb8171d30b16261fab`

## Disponibilità dei file

Questo repository pubblica il **registro delle versioni, gli hash, i changelog e le evidenze di validazione**. Il pacchetto binario completo R4.8 e l'XLSM operativo non sono pubblicati come asset GitHub in questo aggiornamento. Il workbook operativo con dati reali non deve essere pubblicato.

I file `BUILD/`, `BUILD_WEBARCH.cmd` e `PATCH_NOTES_v1.1.1.md` appartengono al builder storico pre-MINIMAL e sono conservati solo per tracciabilità: non sono la procedura STEP10A corrente.

## Documenti correnti

- [Release STEP10A R4.8](releases/beta-step10a-r4.8/RELEASE.md)
- [Identità artefatti R4.8](releases/beta-step10a-r4.8/artifact.json)
- [Verifica R4.8](releases/beta-step10a-r4.8/BUILD_VERIFICATION_R4_8.json)
- [Changeset R4.8](releases/beta-step10a-r4.8/CHANGESET_R4_8.json)
- [Stato beta STEP10A](docs/BETA_STEP10A_R4_8.md)
- [Freeze alpha storico](releases/alpha-1.0.3-home1.1/FREEZE.md)

Conservare l'alpha congelata separatamente. Per ulteriori modifiche beta, partire dalla R4.8 validata e applicare correzioni circoscritte con regressione esplicita.
