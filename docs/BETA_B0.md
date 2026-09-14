# Beta B0 — piano di avvio

Baseline: alpha-1.0.3-home1.1. I 36 componenti produttivi non cambiano nel passaggio di fase. La versione interna VBA resta EASYONE-MINIMAL-1.0.3 per conservare i byte già accettati; la fase beta è documentata separatamente.

## Stato iniziale

L'utente conferma il funzionamento previsto. Alpha archiviata nel pacchetto sorgente completo consegnato nella conversazione; registro del freeze pubblicato in questo repository. Non è stato ricevuto il workbook operativo .xlsm. Non sono stati pubblicati dati reali.

**La sincronizzazione completa dei sorgenti e dell'archivio su GitHub resta da effettuare.** Il ramo beta iniziale contiene il registro, gli hash e questo piano, non una build installabile. I vecchi builder presenti nel repository non vanno usati per ricostruire MINIMAL.

## Lavoro successivo

1. Pubblicare l'archivio esatto identificato in artifact.json e i sorgenti completi, preservando i byte CP1252/CRLF. Verificare hash e copertura, non ricopiare il VBA a mano. Escludere dati reali e workbook operativi.
2. Ripetere su copie di test Excel il ciclo completo: compilazione, importazione, aggiornamento, salvataggio, chiusura e riapertura. Registrare versione Office e report del test, senza dati personali.
3. Eseguire casi riproducibili per stati non mappati, duplicati e conflitti, chiavi mancanti nello staging, date/testi, PDF mascherato XLS, più mensilità, oltre 20 anomalie, HOME ed esportazione.
4. Correggere soltanto difetti riprodotti, una modifica circoscritta per volta. Per ogni modifica indicare componente, motivo, differenziale e regressione; non alterare l'alpha.

## Vincoli conservati

Nessuna reinstallazione della copia funzionante per il solo cambio di fase. Debug del salvataggio mantenuto. Gli stati protocollo non mappati rimangono avvisi non bloccanti; nessuna classificazione canonica inventata. Niente nuove transazioni, rollback, marker o recovery nel percorso ordinario.

La beta non è una dichiarazione di assenza di bug o una release commerciale certificata. È l'avvio del collaudo sistematico dalla baseline funzionante.
