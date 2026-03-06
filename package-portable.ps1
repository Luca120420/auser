# Package Auser Gestione Trasporti Portable
# Creates a portable distribution package with all necessary files

$ErrorActionPreference = "Stop"

Write-Host "Creating portable package for Auser Gestione Trasporti..." -ForegroundColor Cyan

# Define paths
$publishDir = "publish_portable"
$packageDir = "package_temp"
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

# Copy all files from publish_portable
Write-Host "Copying application files..." -ForegroundColor Green
Copy-Item -Path "$publishDir\*" -Destination $packageDir -Recurse -Force

# Create README for portable version
Write-Host "Creating README..." -ForegroundColor Green
$readmeContent = @"
# Auser Gestione Trasporti - Portable Edition

## Descrizione
Questa è la versione portable di Auser Gestione Trasporti. Tutti i dati di configurazione e i contatti dei volontari sono memorizzati nella cartella 'data' all'interno della directory dell'applicazione.

## Requisiti
- Windows 10 o superiore
- .NET 9.0 Runtime (o superiore)

## Installazione
1. Estrarre tutti i file in una cartella a scelta
2. Eseguire AuserGestioneTrasporti.exe

## Caratteristiche Portable
- **Configurazione**: Memorizzata in `data\config.json`
- **Contatti Volontari**: Memorizzati in `data\volunteers.json`
- **Credenziali Gmail**: Criptate con DPAPI (specifiche per macchina)

## Note Importanti
- La cartella 'data' viene creata automaticamente al primo avvio
- Le credenziali Gmail devono essere reinserite quando si copia l'applicazione su un'altra macchina
- Tutti i dati rimangono nella cartella dell'applicazione - nessun file in AppData

## Migrazione da Versione Precedente
Se hai già utilizzato una versione precedente di Auser Gestione Trasporti:
1. Al primo avvio, l'applicazione migrerà automaticamente i dati da AppData
2. La configurazione e i contatti dei volontari saranno copiati nella cartella 'data'
3. Dopo la migrazione, i dati in AppData non saranno più utilizzati

## Struttura Directory
```
AuserGestioneTrasporti/
├── AuserGestioneTrasporti.exe    (Eseguibile principale)
├── data/                          (Dati dell'applicazione)
│   ├── config.json               (Configurazione)
│   └── volunteers.json           (Contatti volontari)
├── README.txt                     (Questo file)
└── [altri file DLL e dipendenze]
```

## Supporto
Per problemi o domande, contattare il team di sviluppo.

---
Versione: Portable Edition
Data Build: $(Get-Date -Format "yyyy-MM-dd")
"@

Set-Content -Path "$packageDir\README.txt" -Value $readmeContent -Encoding UTF8

# Create a placeholder file in data folder to ensure it's included in zip
Write-Host "Creating data folder placeholder..." -ForegroundColor Green
Set-Content -Path "$packageDir\data\.gitkeep" -Value "This folder stores application data" -Encoding UTF8

# Remove test-related files from package (optional cleanup)
Write-Host "Cleaning up test files..." -ForegroundColor Yellow

# Remove test executables and DLLs
Get-ChildItem -Path $packageDir -Filter "testhost.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.TestPlatform.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.VisualStudio.TestPlatform.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.VisualStudio.TraceDataCollector.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "nunit.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "NUnit3.TestAdapter.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Moq.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Castle.Core.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "FsCheck.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "FSharp.Core.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "Microsoft.CodeCoverage.*" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path $packageDir -Filter "testcentric.*" | Remove-Item -Force -ErrorAction SilentlyContinue

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

# Remove localized test resources
$localeDirs = @("cs", "de", "es", "fr", "it", "ja", "ko", "pl", "pt-BR", "ru", "tr", "zh-Hans", "zh-Hant")
foreach ($locale in $localeDirs) {
    $localeDir = "$packageDir\$locale"
    if (Test-Path $localeDir) {
        # Check if directory only contains test-related files
        $files = Get-ChildItem -Path $localeDir -File
        $testFiles = $files | Where-Object { $_.Name -match "(TestPlatform|TraceDataCollector)" }
        if ($files.Count -eq $testFiles.Count) {
            # Only test files, remove entire directory
            Remove-Item -Path $localeDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
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

Write-Host "`nThe portable package includes:" -ForegroundColor Yellow
Write-Host "  - Application executable and all dependencies" -ForegroundColor White
Write-Host "  - Empty 'data' folder for portable storage" -ForegroundColor White
Write-Host "  - README.txt with installation instructions" -ForegroundColor White
Write-Host "  - Automatic migration from AppData on first run" -ForegroundColor White

Write-Host "`nReady to distribute!" -ForegroundColor Green
