================================================================================
  AUSER EXCEL TRANSFORMER - Guida Rapida
================================================================================

REQUISITI DI SISTEMA:
- Windows 10 o Windows 11
- .NET 6.0 Runtime (se non installato, scaricarlo da: https://dotnet.microsoft.com/download/dotnet/6.0)

CONTENUTO DEL PACCHETTO:
- AuserExcelTransformer.exe (applicazione principale)
- File DLL necessari per l'esecuzione
- Cartella "it" con le risorse in italiano

ISTRUZIONI PER L'USO:
1. Estrarre tutti i file dalla cartella zip in una directory a scelta
2. Fare doppio clic su "AuserExcelTransformer.exe" per avviare l'applicazione
3. Selezionare il file CSV con i dati settimanali
4. Selezionare il file Excel esistente (deve contenere il foglio "fissi")
5. Cliccare su "Elabora" per processare i dati
6. Cliccare su "Scarica" per salvare il file Excel aggiornato

STRUTTURA OUTPUT (14 colonne):
1. Data
2. Ora Inizio Servizio
3. Assistito
4. Destinazione
5. Note
6. Auto
7. Volontario
8. Arrivo
9. [vuoto]
10. Indirizzo Partenza
11. Comune Partenza
12. [vuoto]
13. [vuoto]
14. Indirizzo Gasnet

MAPPATURA FOGLIO FISSI:
- Data → Data
- Partenza → Ora Inizio Servizio
- Assistito → Assistito
- Indirizzo → Indirizzo Partenza
- Destinazione → Destinazione
- Note → Note
- Auto → Auto
- Volontario → Volontario
- Arrivo → Arrivo

NOTE:
- Il foglio "fissi" deve essere presente nel file Excel
- I dati con stato "ANNULLATO" vengono automaticamente esclusi
- Le righe con "Accompag. con macchina attrezzata" vengono evidenziate in giallo

SUPPORTO:
Per problemi o domande, contattare l'amministratore di sistema.

Versione: 1.0
Data: Febbraio 2026
================================================================================
