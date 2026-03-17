# Requirements Document

## Introduction

Attualmente, quando l'applicazione genera un nuovo foglio Excel, le colonne "Indirizzo" (colonna 4) e "Note" (colonna 6) vengono popolate con valori statici risolti in C# tramite il `LookupService`, che legge il foglio "assistiti" in memoria. Questa feature richiede che, invece di scrivere valori statici, l'applicazione scriva formule Excel VLOOKUP nelle celle di quelle colonne. Le formule faranno riferimento al foglio "assistiti" presente nella stessa cartella di lavoro, usando il valore della colonna "Assistito" (colonna 3) come chiave di ricerca. In questo modo, se i dati nel foglio "assistiti" vengono aggiornati dopo la generazione del foglio, i valori di "Indirizzo" e "Note" si aggiorneranno automaticamente.

## Glossary

- **ExcelManager**: Il servizio responsabile della scrittura dei dati nel foglio Excel di output.
- **WriteDataRowsEnhanced**: Il metodo di `ExcelManager` che scrive le righe dati nel nuovo foglio generato.
- **Foglio_Assistiti**: Il foglio di riferimento denominato "assistiti" nella cartella di lavoro Excel, con colonne: Nome (col 1), Indirizzo (col 2), Note (col 3).
- **Foglio_Output**: Il nuovo foglio generato dall'applicazione con la struttura a 12 colonne.
- **Colonna_Assistito**: La colonna 3 del Foglio_Output, contenente il nome dell'assistito.
- **Colonna_Indirizzo**: La colonna 4 del Foglio_Output, destinata all'indirizzo dell'assistito.
- **Colonna_Note**: La colonna 6 del Foglio_Output, destinata alle note dell'assistito.
- **VLOOKUP_Formula**: Una formula Excel del tipo `=VLOOKUP(C{row},assistiti!$A:$C,{col_index},0)` che cerca un valore nel Foglio_Assistiti e restituisce il valore dalla colonna specificata.
- **EnhancedTransformedRow**: Il modello dati C# che rappresenta una riga del Foglio_Output.

---

## Requirements

### Requirement 1: Formula VLOOKUP per la colonna Indirizzo

**User Story:** Come utente, voglio che la colonna "Indirizzo" del foglio generato contenga una formula VLOOKUP invece di un valore statico, in modo che l'indirizzo si aggiorni automaticamente se il foglio "assistiti" viene modificato.

#### Acceptance Criteria

1. WHEN `WriteDataRowsEnhanced` scrive una riga nel Foglio_Output, THE ExcelManager SHALL scrivere nella Colonna_Indirizzo una formula VLOOKUP che cerca il valore della Colonna_Assistito della stessa riga nel Foglio_Assistiti e restituisce il valore dalla colonna 2 (Indirizzo).
2. THE ExcelManager SHALL generare la formula nel formato `=VLOOKUP(C{row},assistiti!$A:$C,2,0)` dove `{row}` è il numero di riga corrente nel Foglio_Output.
3. WHEN il nome dell'assistito non è presente nel Foglio_Assistiti, THE ExcelManager SHALL scrivere comunque la formula VLOOKUP nella cella, lasciando la gestione dell'errore a Excel (es. `#N/A`).
4. THE ExcelManager SHALL scrivere la formula come formula Excel (tramite la proprietà `.Formula` della cella EPPlus), non come valore stringa.

### Requirement 2: Formula VLOOKUP per la colonna Note

**User Story:** Come utente, voglio che la colonna "Note" del foglio generato contenga una formula VLOOKUP invece di un valore statico, in modo che le note si aggiornino automaticamente se il foglio "assistiti" viene modificato.

#### Acceptance Criteria

1. WHEN `WriteDataRowsEnhanced` scrive una riga nel Foglio_Output, THE ExcelManager SHALL scrivere nella Colonna_Note una formula VLOOKUP che cerca il valore della Colonna_Assistito della stessa riga nel Foglio_Assistiti e restituisce il valore dalla colonna 3 (Note).
2. THE ExcelManager SHALL generare la formula nel formato `=VLOOKUP(C{row},assistiti!$A:$C,3,0)` dove `{row}` è il numero di riga corrente nel Foglio_Output.
3. WHEN il nome dell'assistito non è presente nel Foglio_Assistiti, THE ExcelManager SHALL scrivere comunque la formula VLOOKUP nella cella, lasciando la gestione dell'errore a Excel (es. `#N/A`).
4. THE ExcelManager SHALL scrivere la formula come formula Excel (tramite la proprietà `.Formula` della cella EPPlus), non come valore stringa.

### Requirement 3: Compatibilità con il flusso di trasformazione esistente

**User Story:** Come sviluppatore, voglio che la modifica alle formule VLOOKUP non alteri il comportamento del resto del flusso di trasformazione, in modo che le altre colonne e funzionalità rimangano invariate.

#### Acceptance Criteria

1. THE ExcelManager SHALL continuare a scrivere valori statici per tutte le colonne del Foglio_Output diverse da Colonna_Indirizzo e Colonna_Note.
2. WHEN `WriteDataRowsEnhanced` viene invocato, THE ExcelManager SHALL preservare il numero di righe scritte, l'ordine delle colonne e il contenuto di tutte le colonne non interessate dalla modifica.
3. THE DataTransformer SHALL continuare a popolare i campi `Indirizzo` e `Note` dell'`EnhancedTransformedRow` tramite `LookupService` (i valori C# rimangono disponibili nel modello, anche se non più scritti come valori statici nel foglio).

### Requirement 4: Correttezza della formula in funzione del numero di riga

**User Story:** Come utente, voglio che ogni riga del foglio generato abbia una formula VLOOKUP che fa riferimento alla propria riga, in modo che ogni assistito ottenga i propri dati corretti.

#### Acceptance Criteria

1. FOR ALL righe scritte da `WriteDataRowsEnhanced`, THE ExcelManager SHALL generare formule VLOOKUP con il numero di riga corretto corrispondente alla riga Excel effettiva in cui la formula viene scritta.
2. WHEN `WriteDataRowsEnhanced` scrive N righe a partire dalla riga `startRow`, THE ExcelManager SHALL generare per la riga i-esima (0-indexed) le formule con riferimento alla riga Excel `startRow + i`.
3. THE ExcelManager SHALL garantire che la formula VLOOKUP nella Colonna_Indirizzo e nella Colonna_Note di ogni riga faccia riferimento alla stessa riga della Colonna_Assistito.
