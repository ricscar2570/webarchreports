# Alpha congelata — 1.0.3 + HOME 1.1

Data del registro: 2026-09-14.

## Decisione

Il responsabile del progetto conferma «funziona tutto come previsto» e richiede il congelamento dell'alpha, l'avvio della beta e l'aggiornamento di GitHub.

La snapshot sorgente è il pacchetto EASYONE MINIMAL 1.0.3 CONSOLIDATO con sostituzione del solo `33_modPulsantiHome.bas` mediante `modPulsantiHome_CORRETTO.bas` HOME 1.1 già consegnato e accettato. Sono 35 moduli standard e una classe. Non è stato cambiato codice produttivo per effettuare questo congelamento. Gli altri 35 componenti e il template sono byte-identici alla 1.0.3.

Restano invariati il debug in modMain, la correzione PA3, gli stati non mappati non bloccanti, le correzioni consolidate e il comportamento dei pulsanti HOME. Nessuna reintroduzione di rollback/recovery nel normale aggiornamento. Le procedure legacy fisicamente presenti non sono state cancellate in questa fase.

## Artefatto completo

Nome: `WebArch_ALPHA_1_0_3_HOME_1_1_CONGELATA.zip`.
Dimensione: 678696 byte.
SHA-256: `e94986fbee71cfcb380fdb53503a91058c812efcc3cbabac03e8c800872ef6f0`.
96 file nel pacchetto, 94 voci nel manifest e 95 checksum, compreso il manifest stesso.

L'archivio è stato creato, riaperto e confrontato con i file verificati. Il pacchetto è consegnato separatamente; **non è stato caricato come asset GitHub e non è incluso nel repository**. Questo commit congela il registro di identità e la decisione di rilascio, non pretende di contenere un albero sorgente completo. Non esiste in questa operazione una GitHub Release binaria.

La copia .xlsm locale funzionante dell'utente non è disponibile nell'ambiente dell'assistente. Non è stato pubblicato alcun dato operativo. Il congelamento dei sorgenti non sostituisce il backup locale di quel workbook.

## Verifiche ripetute

Il wrapper VERIFICA_TUTTO è stato eseguito sulla 1.0.3 + HOME 1.1, in una copia temporanea. Superati i quattro gruppi: regressioni generali, regressioni mirate, contratti delle chiamate produttive, contratti con sorgente dei test nativi. Il pacchetto originale di verifica non è stato modificato dalle prove.

Controlli statici produttivi 10/10; contratti produttivi 120/120; contratti mirati 53/53; 18 + 18 mutazioni negative rilevate; percorsi attivi 4/4. I modelli Python non eseguono VBA.

Un solo contratto lessicale del QA è stato riallineato alla HOME 1.1: riconosce `Address:=""` invece di pretendere la vecchia spelling `Address:=vbNullString`, e verifica la rilettura della destinazione. La correzione riguarda il test, non il modulo produttivo.

L'assistente non ha eseguito Excel o il compilatore VBE. Il collaudo funzionale positivo è dichiarato dall'utente; non attribuiamo automaticamente l'esecuzione di ogni prova nativa inclusa nel pacchetto.

## Avvio beta

Beta B0 utilizza gli stessi hash produttivi elencati in COMPONENTS.sha256. Cambiano fase e documentazione, non le procedure VBA. Nessuna modifica dell'alpha congelata senza una richiesta esplicita; le successive correzioni saranno registrate separatamente.
