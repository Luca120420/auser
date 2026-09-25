# AGENT.md — Auser Gestione Trasporti

> **Read this file before doing anything else on this project.**  
> It contains all the context needed to work correctly: architecture, rules, constraints, naming conventions, build instructions, and known pitfalls.

---

## 1. Project Purpose

Windows desktop app (.NET 9 WinForms) for an Italian volunteer transport organization (AUSER).  
It automates the weekly scheduling workflow:

1. Coordinator receives a **CSV export** from the management system (Gasnet) with service appointments.
2. The app transforms the CSV into a **formatted Excel sheet** appended to a running weekly workbook.
3. The app also manages a **volunteer contact list** and sends **personalized HTML email notifications** to each volunteer via Gmail SMTP.

**Languages used throughout the UI, resources, variable names and comments: Italian.**

---

## 2. Repository Layout

```
auser/
├── AuserExcelTransformer.csproj    # .NET 9 WinForms project (root namespace: AuserExcelTransformer)
├── Program.cs                      # Entry point + DI wiring
├── AGENT.md                        # ← You are here
├── README.md                       # End-user documentation
├── spreadsheet_rules.txt           # 17 original business rules (source of truth for transformation)
├── build_command.txt               # Canonical publish command
│
├── Models/                         # Plain data classes, no logic
├── Services/                       # All business logic
├── UI/                             # WinForms forms, panels, controls, theme
│   └── Controls/                   # Custom WinForms controls
├── Tests/                          # All tests (NUnit + FsCheck + Moq) — NOT compiled into release
├── Properties/                     # Italian resource strings (Resources.resx / Resources.it.resx)
├── Resources/                      # Embedded resources (app_icon.ico)
├── TestData/                       # Sample .xlsx/.csv files used in tests
│
├── DiagnoseTransformation/         # Standalone diagnostic sub-project (excluded from main build)
├── InspectFissi/                   # Standalone inspection tool (excluded from main build)
├── InspectSheets/                  # Standalone inspection tool (excluded from main build)
├── InspectProp/                    # Standalone inspection tool (excluded from main build)
└── Examples/                       # Reference/utility code snippets (excluded from main build)
```

---

## 3. Architecture

The project uses **manual dependency injection** (no DI container). All wiring happens in `Program.cs`.

### Layering

```
UI (MainForm, VolunteerPanel)
    ↓
ApplicationController / VolunteerNotificationController   ← orchestrators
    ↓
Services (ExcelManager, DataTransformer, CSVParser, LookupService, EmailService, …)
    ↓
Models (pure data: ServiceAppointment, EnhancedTransformedRow, Sheet, …)
```

Every service has a corresponding interface (`IExcelManager`, `IDataTransformer`, etc.).  
**Always program to the interface, not the concrete class.**

### Key Services

| Service | File | Responsibility |
|---------|------|---------------|
| `ApplicationController` | `Services/ApplicationController.cs` | Orchestrates the 12-step transform pipeline |
| `ExcelManager` | `Services/ExcelManager.cs` | All Excel I/O (EPPlus), row writing, formatting, sort |
| `DataTransformer` | `Services/DataTransformer.cs` | CSV → `EnhancedTransformedRow` list, applies rules |
| `TransformationRulesEngine` | `Services/TransformationRulesEngine.cs` | Implements the 17 business rules |
| `CSVParser` | `Services/CSVParser.cs` | Parses Italian CSV (CP1252), validates columns |
| `LookupService` | `Services/LookupService.cs` | O(1) VLOOKUP cache for `assistiti` and `fissi` sheets |
| `ColumnStructureManager` | `Services/ColumnStructureManager.cs` | **Single source of truth** for the 12-column output layout |
| `FormattingService` | `Services/FormattingService.cs` | Bold headers, thick borders between date groups |
| `HeaderCalculator` | `Services/HeaderCalculator.cs` | Parses the text header from the previous weekly sheet |
| `DateCalculator` | `Services/DateCalculator.cs` | Date arithmetic (next week same day-of-week) |
| `ConfigurationService` | `Services/ConfigurationService.cs` | JSON persistence for app settings |
| `VolunteerManager` | `Services/VolunteerManager.cs` | Volunteer CRUD, persisted to `volontari-auser.json` |
| `VolunteerNotificationController` | `Services/VolunteerNotificationController.cs` | Email notification workflow |
| `EmailService` | `Services/EmailService.cs` | Gmail SMTP sender (HTML emails) |

---

## 4. Output Excel Structure

The generated sheet has **12 columns**:

| Col | Name | Source |
|-----|------|--------|
| 1 | Data | CSV / fissi / laboratori — format `ddd dd mmm` |
| 2 | Partenza | fissi/laboratori time (h:mm); empty for CSV rows |
| 3 | Assistito | CSV: `CognomeAssistito + " " + NomeAssistito` / fissi / laboratori |
| 4 | Indirizzo | CSV rows: live `VLOOKUP(C{row},assistiti!A:C,2,FALSE)` formula; fissi/laboratori: static copy |
| 5 | Destinazione | CSV: `ComuneDestinazione + IndirizzoDestinazione + CausaleDestinazione` / fissi / laboratori |
| 6 | Note | CSV rows: live `VLOOKUP(C{row},assistiti!A:C,3,FALSE)` formula; fissi/laboratori: static copy |
| 7 | Auto | fissi / laboratori |
| 8 | Volontario | fissi / laboratori |
| 9 | Arrivo | CSV: ORA INIZIO SERVIZIO (h:mm); fissi / laboratori |
| 10 | Avvisi | Lookup on fissi Avv column (CSV rows) / fissi / laboratori |
| 11 | Indirizzo Gasnet | CSV: `ComunePartenza + IndirizzoPartenza` |
| 12 | Note Gasnet | CSV: `NoteERichieste + DescrizionePuntoPartenza` |

**Row 1 (week header):** A1 = Monday date (`dd mmm`, bold 16pt Calibri), B1 = `=A1+6`, C1 = `=CONCATENATE("Settimana ",WEEKNUM(A1))`, D1 = referente placeholder.  
**Row 2:** Column headers (bold, AutoFilter applied).  
**Rows 3+:** Data, sorted by date (col 1) then time (col 2).

### `ColumnStructureManager` is the single source of truth
If you need to add, remove, or rename a column, update `ColumnStructureManager.cs` first, then follow the impact through `ExcelManager`, `DataTransformer`, and `EnhancedTransformedRow`.

---

## 5. Source Sheets in the Excel Workbook

The input workbook **must** contain these sheets:

| Sheet | Required | Purpose |
|-------|----------|---------|
| `fissi` | ✅ Yes | Recurring weekly appointments. 9 columns: Data, Partenza, Assistito, Indirizzo, Destinazione, Note, Auto, Volontario, Arrivo. Dates are shifted +7 days to the same day-of-week in the target week. |
| `assistiti` | ✅ Yes | Reference table for VLOOKUP: col A = full name, col B = Indirizzo, col C = Note. |
| `laboratori` | ❌ Optional | Lab transport appointments. 10 columns (same as fissi + Avv). Same date-shifting logic. |
| `1`, `2`, `3`, … | ✅ At least 1 | Previous weekly output sheets. Numbered 1–53 (never year values like 2025). The app reads the header of the highest-numbered sheet to calculate the next Monday. |

---

## 6. The 12-Step Processing Pipeline

Defined in `ApplicationController.OnProcessButtonClicked()`:

1. Parse CSV → `List<ServiceAppointment>`
2. Get sheet names, compute `nextSheetNumber` (max numeric sheet name + 1, range 1–53)
3. Load `assistiti` and `fissi` sheets
4. Initialize `LookupService` caches
5. `DataTransformer.TransformEnhanced()` → `EnhancedTransformationResult`
6. Read header from sheet `(nextSheetNumber - 1)`, parse with `HeaderCalculator`
7. Compute `nextMondayDate` = previous Monday + 7 days
8. Create new sheet, write header row + column headers, apply column widths + bold
9. `WriteDataRowsEnhanced()` — write CSV rows (VLOOKUP formulas in cols 4 and 6)
10. `AppendFissiData()` — append fissi rows with date-shifting
10.25. `AppendLaboratoriData()` — append laboratori rows (if sheet exists)
10.5. `SortDataRows()` — sort all rows by date then time
10.6. `ApplyThickBordersToDateGroups()` — thick bottom border at end of each date group
12. `AutoFitRowHeights()`, `EnableAutoFilter()`, enable download button

---

## 7. Critical Formatting Rules

These must be preserved across all changes:

1. **Font unification**: `AppendFissiData` and `AppendLaboratoriData` must **not** copy font properties (Bold, Italic, Size, Name, Color) from source cells. Only number formats are copied. This ensures all rows use the worksheet default font. If you see font copying for fissi rows, it is a **bug**.

2. **Yellow highlighting**: Applied inline in `WriteDataRowsEnhanced` for rows where `IsYellow == true`. `SortDataRows` **must preserve** yellow fill after sorting — it does this by capturing `fillColor = "FFFFFF00"` per cell and restoring it after re-writing rows.

3. **Background color**: `AppendFissiData` and `AppendLaboratoriData` must **never** copy the background color from source cells (the fissi sheet often has yellow-highlighted cells that should not bleed into the output).

4. **VLOOKUP formulas**: Columns 4 and 6 of CSV rows must contain live Excel formulas, not resolved string values. `SortDataRows` updates the row numbers in these formulas after sorting using a regex replace on `(?<=[A-Z])\d+`.

5. **Laboratori italic**: `AppendLaboratoriData` applies Italic + Tahoma + size 9 to columns 3 (Assistito) and 4 (Indirizzo) after all other formatting — this overrides any copied font style.

6. **Number formats for dates and times**: 
   - Date column (col 1): `"ddd dd mmm"`
   - Time columns (cols 2 and 9): `"h:mm"`
   - These are always applied explicitly; never rely on the source cell format for these columns.

---

## 8. Business Rules (from `spreadsheet_rules.txt`)

The 17 rules that `TransformationRulesEngine` implements:

- **Rule 1**: Yellow highlight rows where `ATTIVITÀ` contains `"Accompag. con macchina attrezzata"`
- **Rule 3**: Skip rows where `DESCRIZIONE_STATO_SERVIZIO == "ANNULLATO"`
- **Rule 5**: Concatenate `DESCRIZIONE PUNTO PARTENZA` onto `INDIRIZZO PARTENZA`
- **Rules 6–7**: `Assistito` = `CognomeAssistito + " " + NomeAssistito`
- **Rules 10–12**: `Destinazione` = `ComuneDestinazione + IndirizzoDestinazione + CausaleDestinazione`
- **Rule 17**: `IndirizzoGasnet` = `ComunePartenza + IndirizzoPartenza`
- **Note Gasnet**: `NoteERichieste + DescrizionePuntoPartenza` (concatenated with space)

Always consult `spreadsheet_rules.txt` for the authoritative rule definitions.

---

## 9. CSV Input Format

Parsed by `CSVParser` using **CsvHelper** with Italian encoding (**CP1252** via `System.Text.Encoding.CodePages`).

Required columns (validated before parsing):
- `DATA SERVIZIO`
- `ORA INIZIO SERVIZIO`
- `COGNOME ASSISTITO`
- `NOME ASSISTITO`

All 13 columns:
`DATA SERVIZIO, ORA INIZIO SERVIZIO, ATTIVITÀ, DESCRIZIONE STATO SERVIZIO, INDIRIZZO PARTENZA, COMUNE PARTENZA, DESCRIZIONE PUNTO PARTENZA, INDIRIZZO DESTINAZIONE, COMUNE DESTINAZIONE, CAUSALE DESTINAZIONE, COGNOME ASSISTITO, NOME ASSISTITO, NOTE E RICHIESTE`

The `ServiceAppointmentMap` class inside `CSVParser.cs` maps these Italian headers to C# property names.

---

## 10. Build, Test, Publish

### Build
```bash
dotnet build
```

### Run tests
```bash
dotnet test
```

Tests are in `Tests/`. The `AuserExcelTransformer.csproj` explicitly includes only the test files listed under `<Compile Include="Tests\...">` — tests not listed there are excluded from compilation.

Test frameworks: **NUnit** (unit tests), **FsCheck** (property-based), **Moq** (mocking).  
These packages have `<PrivateAssets>all</PrivateAssets>` and are excluded from publish output.

### Publish (canonical command — do not change)
```bash
dotnet publish AuserExcelTransformer.csproj -c Release -r win-x64 --self-contained true -o build_output
```

Post-publish MSBuild target (`PatchSubsystem`) automatically:
1. Patches the PE header subsystem byte from `3` (console) to `2` (GUI) to suppress the console window
2. Deletes everything in `build_output/` except `auser_gestione_trasporti_v2.0.2.exe`

Result: **single-file, self-contained Windows x64 executable** (~60–80 MB).

---

## 11. NuGet Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| EPPlus | 7.0.5 | Excel `.xlsx` read/write |
| CsvHelper | 30.0.1 | CSV parsing |
| System.Text.Encoding.CodePages | 10.0.4 | CP1252 encoding (Italian CSV) |
| Newtonsoft.Json | 13.0.1 | App config + volunteer JSON |
| NUnit | 4.0.1 | Unit testing (test-only) |
| FsCheck | 2.16.6 | Property-based testing (test-only) |
| Moq | 4.20.70 | Mocking (test-only) |
| Microsoft.NET.Test.Sdk | 17.8.0 | Test runner (test-only) |

---

## 12. Project Configuration

- **Target framework**: `net9.0-windows`
- **Root namespace**: `AuserExcelTransformer`
- **Assembly name**: `auser_gestione_trasporti_v2.0.2`
- **Output type**: `WinExe`
- **Implicit usings**: **disabled** — every `using` must be explicit
- **Nullable**: enabled
- **Localization**: Italian (`SatelliteResourceLanguages=it`). All user-facing strings must go through `Properties/Resources.resx`; never hardcode UI text.
- **EPPlus license**: `LicenseContext.NonCommercial` — set in `ExcelManager.OpenWorkbook`. Do not remove.

---

## 13. UI Structure

Built with **WinForms**, fully programmatic layout (no `.Designer.cs` for the main layout logic).

### Theme (`UI/ThemeManager.cs`)
| Token | Value | Usage |
|-------|-------|-------|
| `ColorAppBackground` | `#edf6f0` | Form and panel backgrounds |
| `ColorPrimary` | `#068534` | Accent color, primary buttons, active nav |
| `ColorSecondary` | mid-gray | Subtitles, inactive nav labels |
| `ColorBorderLight` | light gray | Separator lines |
| `ColorSoft` | light mint | Active nav pill background |
| `ColorError` | `#D32F2F` | Error messages |

Font: **Segoe UI** throughout.

### Layout
- Fixed 280px **sidebar**: app title + subtitle + 2 nav items (pills)
- **PageHeader** (64px): current page title
- **ContentPanel** (scrollable, fills rest): centered `InnerPanel`, max width 1080px, responsive

### Pages
1. **Aggiungi Accompagnamenti** (`_transformPage`): single `RoundedPanel` card with CSV/Excel file selectors, Elabora + Salva action buttons, status label.
2. **Gestione Volontari** (`VolunteerPanel`): volunteer `ListView`, Gmail credentials, Excel file selector, email send with progress.

### Custom Controls
- `ModernButton` (`UI/Controls/ModernButton.cs`): custom-drawn, variants via `ThemeManager.ApplyPrimary()` / `ApplyOutlinePrimary()`
- `RoundedPanel` (`UI/Controls/RoundedPanel.cs`): panel with rounded corners and optional accent bar

---

## 14. Data Persistence

| Data | File | Format |
|------|------|--------|
| App configuration (Gmail credentials, last paths) | `AppConfiguration` via `ConfigurationService` | JSON |
| Volunteer contacts | `volontari-auser.json` | JSON `{ "associates": { "Surname": "email" } }` |

---

## 15. Known Pitfalls and Non-Obvious Decisions

1. **Sheet numbering range is 1–53**: `GetNextSheetNumber` ignores numeric sheet names outside this range to avoid mistaking year sheets (e.g., `"2025"`) for weekly sheet numbers.

2. **`GetSheetByName` is whitespace-tolerant**: Trims both the search name and sheet names before comparing. Always use `GetSheetByName` or `GetFissiSheet` — never access `Package.Workbook.Worksheets` directly.

3. **Fissi data start row detection**: `AppendFissiData` detects whether the fissi sheet uses row 1 or row 2 as column headers by checking if cell A2 contains "Data". Default start row is 3.

4. **Time value normalization**: Time columns in fissi/laboratori may arrive as `double` (Excel serial), `DateTime`, or `string` (e.g., `"8.30.00"` with dots). `AppendFissiData` and `AppendLaboratoriData` normalize all of these to an Excel time fraction (`TimeSpan.TotalDays`).

5. **SortDataRows regex for formula row updates**: Uses `(?<=[A-Z])\d+` to update row numbers inside formulas (e.g., `C3 → C5`). This is intentionally simple — it works because VLOOKUP formulas in this sheet only reference the current row, not ranges.

6. **Post-sort white fill**: `SortDataRows` explicitly sets `ExcelFillStyle.Solid` with white for non-yellow cells when writing back rows. This is intentional — EPPlus can bleed shared style objects if a fill is not explicitly set, causing non-yellow cells to appear yellow.

7. **`ImplicitUsings` is disabled**: Every file needs its own `using` statements. Do not rely on global usings.

8. **Tests not auto-compiled**: Only the `<Compile Include="Tests\...">` entries in the csproj are compiled. If you add a new test file, you must also add it to the csproj include list.

9. **Localized resource strings**: All user-visible strings (error messages, button labels, success messages) are in `Properties/Resources.resx`. Reference them as `Resources.SomeKey`. Never hardcode UI text in C# files.

10. **`CSVPreprocessor`** exists to handle malformed CSV lines before `CsvHelper` processes them. If adding CSV parsing logic, check whether it belongs there or in `CSVParser`.

11. **EmailService uses App Passwords**: Gmail SMTP requires an App Password (not the Gmail account password). Credentials are stored via `GmailCredentials` in `AppConfiguration`.

---

## 16. Adding New Features — Checklist

- [ ] Does it change the output column structure? → Update `ColumnStructureManager` first.
- [ ] Does it add a new column? → Update `EnhancedTransformedRow`, `DataTransformer`, `WriteDataRowsEnhanced`, `AppendFissiData`, `AppendLaboratoriData`, `SortDataRows` (add capture + restore for new formatting).
- [ ] Does it add a new source sheet? → Follow the pattern of `AppendFissiData` / `AppendLaboratoriData`.
- [ ] Does it touch formatting? → Review the 7 critical formatting rules in section 7.
- [ ] Does it add a user-facing string? → Add to `Properties/Resources.resx`, reference via `Resources.Key`.
- [ ] Does it add a new test file? → Add the `<Compile Include="Tests\NewFile.cs" />` entry to `AuserExcelTransformer.csproj`.
- [ ] After any code change → Run `dotnet build` to verify compilation.

---

## 17. Interface Contracts Summary

```csharp
// Core interfaces — always inject these, never the concrete class
IApplicationController   // OnCSVFileSelected, OnExcelFileSelected, OnProcessButtonClicked, OnDownloadButtonClicked
IExcelManager            // OpenWorkbook, GetFissiSheet, GetSheetByName, CreateNewSheet, WriteDataRowsEnhanced,
                         // AppendFissiData, AppendLaboratoriData, SortDataRows, ApplyThickBordersToDateGroups,
                         // WriteHeaderRow, WriteColumnHeadersEnhanced, ApplyColumnWidths, AutoFitRowHeights,
                         // ApplyBoldToHeaders, EnableAutoFilter, SaveWorkbook, GetSheetNames, GetNextSheetNumber,
                         // ReadHeader, ApplyYellowHighlight
IDataTransformer         // Transform(appointments), TransformEnhanced(appointments, lookupService)
ICSVParser               // ParseCSV(path), ValidateCSVStructure(path, out missingColumns)
IHeaderCalculator        // ParseHeader(headerText) → HeaderInfo
IDateCalculator          // CalculateNextWeekSameDay(sourceDate, targetMonday)
ILookupService           // LoadReferenceSheets(assistiti, fissi), LookupInAssistiti(name, col), LookupInFissi(name, col)
IFormattingService       // ApplyBoldHeaders(sheet, headerRow), ApplyDateGroupBorders(sheet, start, end, col)
IVolunteerManager        // LoadVolunteers, SaveVolunteers, AddVolunteer, RemoveVolunteer, IsValidEmail
IEmailService            // SendVolunteerNotificationAsync(credentials, volunteer, rows)
IConfigurationService    // LoadConfiguration, SaveConfiguration
IGUI                     // ShowWindow, ShowErrorMessage, ShowSuccessMessage, DisplaySelectedCSVPath,
                         // DisplaySelectedExcelPath, EnableProcessButton, EnableDownloadButton, GetSaveFilePath
```
