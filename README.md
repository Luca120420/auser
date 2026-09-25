# Auser Gestione Trasporti

Applicazione desktop Windows che automatizza la preparazione settimanale dei turni di trasporto per i volontari di AUSER. Ogni settimana il coordinatore riceve un file CSV dal gestionale (es. Gasnet) con gli appuntamenti di servizio, e l'app lo trasforma in un foglio Excel formattato aggiunto al registro storico. Gestisce anche la rubrica dei volontari e l'invio di email di notifica personalizzate.

**In sintesi: CSV in → foglio settimanale Excel formattato + notifiche email ai volontari.**

---

## Funzionalità principali

- **Importazione CSV**: legge l'export del gestionale (13 colonne, encoding CP1252/italiano) e applica le 17 regole di trasformazione dati.
- **Fogli fissi e laboratori**: aggiunge automaticamente le righe ricorrenti dai fogli `fissi` e `laboratori` (opzionale) con scorrimento automatico delle date alla settimana successiva.
- **Lookup VLOOKUP**: colonne Indirizzo e Note per le righe CSV sono scritte come formule `VLOOKUP` live su foglio `assistiti`; i valori rimangono aggiornati se la rubrica cambia.
- **Ordinamento**: tutte le righe (CSV, fissi, laboratori) vengono ordinate per data poi ora, con conservazione di formule, colori e formati.
- **Evidenziazione gialla**: le righe con "Accompag. con macchina attrezzata" vengono evidenziate in giallo, con il colore preservato dopo l'ordinamento.
- **Formattazione coerente**: bordi sottili su tutte le celle, bordo spesso a separare i gruppi per data, testo a capo nelle colonne di testo, altezze riga auto-adattate. Il font dei fogli sorgente non viene copiato: tutte le righe ereditano il font di default del foglio.
- **AutoFilter**: abilitato sulla riga delle intestazioni di colonna.
- **Gestione volontari**: rubrica dei volontari (nome → email) salvata in JSON, importabile/esportabile.
- **Invio email Gmail**: email HTML personalizzata per ogni volontario con i propri turni assegnati, via SMTP Gmail (TLS porta 587).

---

## Struttura dell'output Excel

Il foglio generato ha 12 colonne:

| Col | Nome | Fonte |
|-----|------|-------|
| 1 | Data | CSV / fissi / laboratori (formato `ddd dd mmm`) |
| 2 | Partenza | Fissi/laboratori (h:mm); vuoto per le righe CSV |
| 3 | Assistito | CSV (`Cognome Nome`) / fissi / laboratori |
| 4 | Indirizzo | Formula `VLOOKUP` su `assistiti` (righe CSV) / copia diretta (fissi, laboratori) |
| 5 | Destinazione | CSV (`Comune + Indirizzo + Causale`) / fissi / laboratori |
| 6 | Note | Formula `VLOOKUP` su `assistiti` (righe CSV) / copia diretta (fissi, laboratori) |
| 7 | Auto | Fissi / laboratori |
| 8 | Volontario | Fissi / laboratori |
| 9 | Arrivo | ORA INIZIO SERVIZIO dal CSV (h:mm) / fissi / laboratori |
| 10 | Avvisi | Lookup su colonna Avv del foglio fissi (righe CSV) |
| 11 | Indirizzo Gasnet | CSV (`Comune Partenza + Indirizzo Partenza`) |
| 12 | Note Gasnet | CSV (`Note e Richieste + Descrizione Punto Partenza`) |

**Riga 1** (intestazione settimana): data lunedì (`A1`), formula domenica (`=A1+6`), numero settimana (`CONCATENATE("Settimana ", WEEKNUM(A1))`), campo referente.  
**Riga 2**: intestazioni di colonna in grassetto con AutoFilter.  
**Righe 3+**: dati, ordinati per data e ora.

---

## Struttura del progetto

```
auser/
├── AuserExcelTransformer.csproj      # Progetto .NET 9 WinForms
├── Program.cs                        # Entry point + dependency injection
├── spreadsheet_rules.txt             # Documentazione delle 17 regole di trasformazione
├── build_command.txt                 # Comando di pubblicazione di riferimento
│
├── Models/
│   ├── ServiceAppointment.cs         # 13 campi dal CSV
│   ├── EnhancedTransformedRow.cs     # 12 colonne di output + flag IsYellow
│   ├── TransformedRow.cs             # Modello legacy
│   ├── EnhancedTransformationResult.cs
│   ├── TransformationResult.cs
│   ├── HeaderInfo.cs                 # Intestazione settimana analizzata
│   ├── ExcelWorkbook.cs              # Wrapper EPPlus ExcelPackage
│   ├── Sheet.cs                      # Wrapper EPPlus ExcelWorksheet
│   ├── AppConfiguration.cs           # Impostazioni persistite
│   ├── GmailCredentials.cs
│   ├── VolunteerAssignment.cs
│   └── VolunteerFileData.cs
│
├── Services/
│   ├── ApplicationController.cs      # Orchestratore principale (12 step)
│   ├── ExcelManager.cs               # I/O Excel completo (EPPlus)
│   ├── DataTransformer.cs            # CSV → righe di output
│   ├── TransformationRulesEngine.cs  # 17 regole di trasformazione
│   ├── CSVParser.cs                  # Parser CSV (CsvHelper, CP1252)
│   ├── CSVPreprocessor.cs
│   ├── LookupService.cs              # Lookup O(1) su assistiti/fissi
│   ├── ColumnStructureManager.cs     # Definizione 12 colonne
│   ├── FormattingService.cs          # Intestazioni grassetto + bordi gruppi data
│   ├── HeaderCalculator.cs           # Analisi intestazione foglio precedente
│   ├── DateCalculator.cs             # Utilità calcolo date
│   ├── ConfigurationService.cs       # Persistenza configurazione JSON
│   ├── VolunteerManager.cs           # CRUD rubrica volontari + JSON
│   ├── VolunteerNotificationController.cs # Workflow notifiche email
│   └── EmailService.cs               # Invio SMTP Gmail (HTML)
│
├── UI/
│   ├── MainForm.cs                   # Finestra principale (layout programmatico)
│   ├── VolunteerPanel.cs             # Pagina gestione volontari
│   ├── ThemeManager.cs               # Palette colori + stili pulsanti
│   ├── IGUI.cs                       # Interfaccia astrazione GUI
│   └── Controls/
│       ├── ModernButton.cs           # Pulsante personalizzato (Primary/Outline/ecc.)
│       └── RoundedPanel.cs           # Panel con angoli arrotondati
│
├── Tests/                            # Test unitari e property-based (NUnit + FsCheck + Moq)
├── Properties/                       # Stringhe UI italiane (Resources.resx / Resources.it.resx)
├── Resources/                        # Icona app (app_icon.ico)
├── TestData/                         # File .xlsx/.csv di esempio per i test
└── logo/                             # Asset PNG del logo
```

---

## Prerequisiti

- Windows 10 o Windows 11 (x64)
- [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9) per compilare da sorgente
- Non è richiesta l'installazione di Microsoft Excel

---

## Build e pubblicazione

Build standard:
```bash
dotnet build
```

Esecuzione test:
```bash
dotnet test
```

**Pubblicazione (eseguibile autonomo, single-file per Windows x64):**
```bash
dotnet publish AuserExcelTransformer.csproj -c Release -r win-x64 --self-contained true -o build_output
```

Il processo di pubblicazione:
1. Compila il progetto in Release
2. Applica un patch al PE header per sopprimere la finestra console
3. Rimuove tutti i file tranne `auser_gestione_trasporti_v2.0.2.exe` dalla cartella `build_output`

Il risultato è un singolo file `.exe` autonomo (~60–80 MB).

---

## Dipendenze NuGet

| Pacchetto | Versione | Utilizzo |
|-----------|----------|---------|
| EPPlus | 7.0.5 | Lettura/scrittura file `.xlsx` |
| CsvHelper | 30.0.1 | Parsing CSV |
| System.Text.Encoding.CodePages | 10.0.4 | Encoding CP1252 (italiano) |
| Newtonsoft.Json | 13.0.1 | Configurazione e rubrica volontari JSON |
| NUnit | 4.0.1 | Framework di test unitari |
| FsCheck | 2.16.6 | Test property-based |
| Moq | 4.20.70 | Mocking per i test |

---

## Interfaccia utente

L'applicazione ha due sezioni accessibili dalla barra di navigazione laterale:

**⇄ Aggiungi Accompagnamenti**  
Seleziona il file CSV e il file Excel (che deve contenere i fogli `fissi` e `assistiti`), premi *Elabora* per generare il nuovo foglio settimanale, poi *Salva* per esportare il file modificato.

**👥 Gestione Volontari**  
Gestisci la rubrica email dei volontari, configura le credenziali Gmail (email + app password), seleziona il foglio Excel e il numero di settimana, e invia le notifiche email personalizzate a ogni volontario con i propri turni assegnati.

---

## Licenza

Applicazione sviluppata per AUSER, organizzazione di volontariato italiana.
