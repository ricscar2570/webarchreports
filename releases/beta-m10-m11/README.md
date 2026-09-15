# WebArch Beta — M10 + M11 gate candidate

Data: 15 settembre 2026.

## Stato

- Baseline preservata: `alpha-1.0.3-home1.1` / `EASYONE-MINIMAL-1.0.3` + HOME 1.1.
- M10 implementato come estensione additiva: comandi unificati, `OP_HOME`, `OP_SINTESI`, `OP_REPORT`, controllo di staleness e report operativo separato.
- QA statico M10: **108/108 PASS**.
- Model tests M10: **12/12 PASS**.
- M11: **gate candidate**. Non è ancora una RC congelata.
- Restano obbligatori su Excel desktop: `Debug > Compila VBAProject`, test nativi, ciclo completo Save/chiusura/riapertura e misura sui dati reali.

## Sorgenti pubblicati

`src/M10/` contiene i quattro moduli produttivi M10 e il test nativo M10.
`src/M11/` contiene il test di accettazione M11.

I moduli M10/M11 sono aggiuntivi: nessun componente alpha, FSC B1 o M1–M9 viene sostituito da questa directory.

## Archivio consegnato

Archivio completo conservato nella Library del progetto:

`WebArch_BETA_M10_M11_CONSEGNA_FINAL_20260915.zip`

SHA-256:

`9edb1a64fab55d19cf6486191c717fe7f3c59f697fc55c8c306dc1316c51fe76`

Il manifest sottostante consente di verificare i singoli sorgenti. Gli ZIP binari non vengono ricostruiti dal repository: fanno fede l'archivio conservato nella Library e il relativo checksum.

## Regola di rilascio

Lo stato corretto fino al completamento del gate nativo è:

**M11 CANDIDATO — GATE NATIVO NON ESEGUITO**

Solo dopo PASS reale su Excel, Save/reopen e approvazione utente si potrà congelare una Release Candidate.
