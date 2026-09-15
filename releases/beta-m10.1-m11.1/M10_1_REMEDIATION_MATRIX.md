# M10.1 Audit Remediation — matrice dei difetti

| ID audit | Stato M10.1 | Correzione |
|---|---|---|
| P0-01 | IMPLEMENTATO | Eliminato il richiamo dei pulsanti RICARICA; M10.1 usa le API core M3/M4/M5/M7/M8/M9. Test nativo sotto BUSY incluso. |
| P0-02 | IMPLEMENTATO | LAST_GOOD separato da LAST_ATTEMPT; il fallimento non modifica LAST_GOOD. |
| P0-03 | IMPLEMENTATO | `MINIMAL_DiscardRequired` blocca IsCurrent, rebuild, riallineamento ed export. |
| P1-01 | IMPLEMENTATO | Configurazione letta tramite `SH_CONFIG`; nessuna euristica CONFIG/RunID. |
| P1-02 | IMPLEMENTATO | FP2 usa serializzazione completa di ogni valore/tipo; rimosso TextAtom campionato. |
| P1-03 | IMPLEMENTATO | Fingerprint valido solo con prefisso FP2; ERR/MISSING/vuoto rifiutati. |
| P1-04 | IMPLEMENTATO CONSERVATIVO | Token include l'intero `SH_CONFIG` e i fogli config/map/defs rilevanti. Può produrre falsi stale, mai falso current. |
| P1-05 | IMPLEMENTATO | Protocolli, Protocol Plus e Scadenze sono refresh obbligatori; sync M5 precede refresh M5. |
| P1-06 | IMPLEMENTATO | L'esito delle API core deve essere Boolean/numerico verificabile; nessun OnAction. |
| P1-07 | IMPLEMENTATO FAIL-CLOSED | B1 conserva `FSC_Aggiorna`, ma M10.1 confronta l'intero stato `FSC_*`, controlla discard e salva solo dopo riallineamento completo. Il gate nativo deve confermare il contratto B1 sul workbook reale. |
| P1-08 | IMPLEMENTATO FAIL-CLOSED | Preflight interroga i contatori visibili M3/M4/M5/M7/M8/M9; API assente = blocco prima di scrivere. |
| P1-09 | IMPLEMENTATO CONSERVATIVO | Nessuna scansione testuale UsedRange per flags: segnalazioni = N/D se non esiste API strutturata. |
| P1-10 | IMPLEMENTATO | Conteggio viste via ListObject; in assenza di tabella = N/D, non conteggio inventato. |
| P1-11 | IMPLEMENTATO | Last row e last column trovati con `Find`, non con `UsedRange`. |
| P2-01 | HARDENING M4 SEPARATO | Non modifica M10.1. Le due correzioni M4 sono descritte in `M4H1_HARDENING_REQUIRED.txt` e restano requisito prima di RC1. |
| P2-02 | IMPLEMENTATO | Guard `Nothing` sequenziale in M10.1. |
| P2-03 | IMPLEMENTATO | Installer traccia i fogli preesistenti e rimuove soltanto quelli M10 creati nel tentativo fallito. |
| P2-04 | IMPLEMENTATO | Test M10.1 e M11.1 catturano Err.Number/Description prima del cleanup. |
| P2-05 | IMPLEMENTATO | M11.1 richiede FP validi, discard=false e prova reale del riallineamento sotto BUSY. |
| P2-06 | IMPLEMENTATO CONSERVATIVO | M4 viene ricaricato; i valori custom M4 non sono esportati nel PDF M10.1, quindi IncludeExport non viene aggirato. |
| PKG-01 | NUOVO PACKAGING CORRETTO | Manifest M10.1/M11.1 generato solo a fine QA. Gli storici M5-M9 restano invariati. |
