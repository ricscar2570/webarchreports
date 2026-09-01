# WebArch v1.1.1 — Builder hotfix

Correzione del builder Windows dopo il primo test reale su Excel del 1 settembre 2026.

## Difetto rilevato
Il template `.xlsx` poteva esporre `Workbook.VBProject` ma restituire `VBProject.VBComponents = $null`. PowerShell interpreta `$null.Count` come `0`, quindi il vecchio log mostrava `Componenti iniziali: 0` e il primo `VBComponents.Import()` falliva con `Impossibile chiamare un metodo su un'espressione con valore null.`

## Correzione
Il builder ora converte prima il template in `.xlsm`, chiude e riapre il workbook, verifica esplicitamente `VBProject` e `VBComponents`, quindi importa i moduli. In caso di errore registra anche riga PowerShell, comando e stack.

La patch corregge il difetto osservato in runtime; il collaudo completo Excel/Windows resta necessario prima della classificazione RC2/produzione.
