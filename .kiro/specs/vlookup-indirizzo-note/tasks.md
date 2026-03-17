# Tasks

## Task List

- [x] 1. Modificare WriteDataRowsEnhanced in ExcelManager
  - [x] 1.1 Sostituire la scrittura del valore statico di Indirizzo (col 4) con la formula `.Formula = $"VLOOKUP(C{currentRow},assistiti!$A:$C,2,0)"`
  - [x] 1.2 Sostituire la scrittura del valore statico di Note (col 6) con la formula `.Formula = $"VLOOKUP(C{currentRow},assistiti!$A:$C,3,0)"`

- [x] 2. Scrivere i test per la feature
  - [x] 2.1 Unit test: singola riga con startRow=2 — verificare formula esatta in col 4 e col 6
  - [x] 2.2 Unit test: verificare che le colonne non interessate (1,2,3,5,7-12) abbiano Formula vuota
  - [x] 2.3 Unit test: lista vuota — nessuna eccezione
  - [x] 2.4 [PBT] Property test: per qualsiasi lista di righe e startRow, col 4 e col 6 contengono una formula VLOOKUP (Property 1)
  - [x] 2.5 [PBT] Property test: per qualsiasi lista di righe e startRow, la formula nella riga i-esima referenzia la riga Excel startRow+i (Property 2)
  - [x] 2.6 [PBT] Property test: per qualsiasi lista di righe, le colonne diverse da 4 e 6 hanno Formula vuota (Property 3)

- [x] 3. Verificare la build
  - [x] 3.1 Eseguire `dotnet publish AuserExcelTransformer.csproj -c Release -r win-x64 --self-contained true -o build_output` e confermare che non ci siano errori di compilazione
