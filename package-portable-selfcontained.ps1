# Package Auser Gestione Trasporti Portable (Self-Contained)
# Creates a portable distribution package with .NET runtime included

$ErrorActionPreference = "Stop"

Write-Host "Creating self-contained portable package for Auser Gestione Trasporti..." -ForegroundColor Cyan

# Define paths
$publishDir = "publish_portable_selfcontained"
$packageDir = "package_temp_selfcontained"
$zipFile = "auser_gestione_trasporti_portable.zip"

# Clean up any existing package directory
if (Test-Path $packageDir) {
    Write-Host "Cleaning up existing package directory..." -ForegroundColor Yellow
    Remove-Item -Path $packageDir -Recurse -Force
}

# Create package directory structure
Write-Host "Creating package directory structure..." -ForegroundColor Green
New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
New-Item -ItemType Directory -Path "$packageDir\data" -Force | Out-Null

# Copy all files from publish_portable_selfcontained
Write-Host "Copying application files..." -ForegroundColor Green
Copy-Item -Path "$publishDir\*" -Destination $packageDir -Recurse -Force

# Create README for portable version
Write-Host "Creating README..." -ForegroundColor Green
$readmeContent = @"
# Auser Gestione Trasporti - Portable Edition (Self-Contained)

## Descrizione
Questa è la versione portable SELF-CONTAINED di Auser Gestione Trasporti. 
Include il runtime .NET 9.0 - NON è necessario installare nulla!

Tutti i dati di configurazione e i contatti dei volontari sono memorizzati nella cartella 'data' all'interno della directory dell'applicazione.

## Requisiti
- Windows 10 o superiore
- NESSUN ALTRO REQUISITO - il runtime .NET è incluso!

## Installazione
1. Estrarre tutti i file in una cartella a scelta
2. Eseguire AuserGestioneTrasporti.exe
3. Fatto! L'applicazione funziona immediatamente

## Caratteristiche Portable
- **Configurazione**: Memorizzata in ``data\config.json``
- **Contatti Volontari**: Memorizzati in ``data\volunteers.json``
- **Credenziali Gmail**: Criptate con DPAPI (specifiche per macchina)
- **Runtime .NET**: Incluso - nessuna installazione richiesta!

## Note Importanti
- La cartella 'data' viene creata automaticamente al primo avvio
- Le credenziali Gmail devono essere reinserite quando si copia l'applicazione su un'altra macchina
- Tutti i dati rimangono nella cartella dell'applicazione - nessun file in AppData
- Questa versione è più grande (~60 MB) perché include il runtime .NET

## Migrazione da Versione Precedente
Se hai già utilizzato una versione precedente di Auser Gestione Trasporti:
1. Al primo avvio, l'applicazione migrerà automaticamente i dati da AppData
2. La configurazione e i contatti dei volontari saranno copiati nella cartella 'data'
3. Dopo la migrazione, i dati in AppData non saranno più utilizzati

## Struttura Directory
``````
AuserGestioneTrasporti/
├── AuserGestioneTrasporti.exe    (Eseguibile principale)
├── data/                          (Dati dell'applicazione)
│   ├── config.json               (Configurazione)
│   └── volunteers.json           (Contatti volontari)
├── README.txt                     (Questo file)
└── [runtime .NET e dipendenze]
``````

## Risoluzione Problemi

### L'applicazione non si avvia
- Verificare di aver estratto TUTTI i file dalla cartella zip
- Verificare che Windows non abbia bloccato i file (click destro > Proprietà > Sblocca)
- Eseguire come amministratore se necessario

### Errore "Impossibile trovare il runtime"
- Questa versione include il runtime - se vedi questo errore, alcuni file potrebbero essere mancanti
- Riestrarre tutti i file dalla cartella zip

## Supporto
Per problemi o domande, contattare il team di sviluppo.

---
Versione: Portable Self-Contained Edition
Runtime: .NET 9.0 (incluso)
Data Build: $(Get-Date -Format "yyyy-MM-dd")
"@

Set-Content -Path "$packageDir\README.txt" -Value $readmeContent -Encoding UTF8

# Create empty config and volunteers files with proper structure
Write-Host "Creating empty data files..." -ForegroundColor Green

$emptyConfig = @"
{
  "GmailEmail": "",
  "GmailAppPassword": "",
  "VolunteerFilePath": ""
}
"@

$emptyVolunteers = @"
{
  "associates": {}
}
"@

Set-Content -Path "$packageDir\data\config.json" -Value $emptyConfig -Encoding UTF8
Set-Content -Path "$packageDir\data\volunteers.json" -Value $emptyVolunteers -Encoding UTF8

# Remove test-related files from package
Write-Host "Cleaning up test files..." -ForegroundColor Yellow

# Remove test executables and DLLs
Get-ChildItem -Path $packageDir -Filter "testhost.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.TestPlatform.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.VisualStudio.TestPlatform.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.VisualStudio.TraceDataCollector.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "nunit.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "NUnit3.TestAdapter.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Moq.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Castle.Core.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "FsCheck.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "FSharp.Core.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.CodeCoverage.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "testcentric.*" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue

# Remove test-related directories
if (Test-Path "$packageDir\CodeCoverage") {
    Remove-Item -Path "$packageDir\CodeCoverage" -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path "$packageDir\TestData") {
    Remove-Item -Path "$packageDir\TestData" -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path "$packageDir\InstrumentationEngine") {
    Remove-Item -Path "$packageDir\InstrumentationEngine" -Recurse -Force -ErrorAction SilentlyContinue
}

# Create zip file
Write-Host "Creating zip archive..." -ForegroundColor Green
if (Test-Path $zipFile) {
    Remove-Item -Path $zipFile -Force
}

Compress-Archive -Path "$packageDir\*" -DestinationPath $zipFile -CompressionLevel Optimal

# Clean up temporary package directory
Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
Remove-Item -Path $packageDir -Recurse -Force

# Display results
Write-Host "`nPackage created successfully!" -ForegroundColor Green
Write-Host "File: $zipFile" -ForegroundColor Cyan
$zipSize = (Get-Item $zipFile).Length / 1MB
Write-Host "Size: $([math]::Round($zipSize, 2)) MB" -ForegroundColor Cyan

Write-Host "`nThe self-contained portable package includes:" -ForegroundColor Yellow
Write-Host "  - Application executable and all dependencies" -ForegroundColor White
Write-Host "  - .NET 9.0 Runtime (NO installation required!)" -ForegroundColor White
Write-Host "  - Empty 'data' folder with config.json and volunteers.json" -ForegroundColor White
Write-Host "  - README.txt with installation instructions" -ForegroundColor White
Write-Host "  - Automatic migration from AppData on first run" -ForegroundColor White

Write-Host "`nReady to distribute - works on ANY Windows 10+ computer!" -ForegroundColor Green
