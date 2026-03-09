# Test script to verify Destinazione column includes CAUSALE DESTINAZIONE

Write-Host "Testing Destinazione Column with CAUSALE DESTINAZIONE..." -ForegroundColor Cyan
Write-Host ""

# Build the project
Write-Host "Building project..." -ForegroundColor Yellow
dotnet build AuserExcelTransformer.csproj -c Debug -v quiet

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed!" -ForegroundColor Red
    exit 1
}

Write-Host "Build successful!" -ForegroundColor Green
Write-Host ""

# Run the transformation
$csvFile = "168514-Estrazione_1772702766454.csv"
$outputFile = "test_output_destinazione.xlsx"

Write-Host "Processing CSV file: $csvFile" -ForegroundColor Yellow
Write-Host "Output file: $outputFile" -ForegroundColor Yellow
Write-Host ""

# Remove old output if exists
if (Test-Path $outputFile) {
    Remove-Item $outputFile
}

# Run the application
$exePath = "bin\Debug\net9.0-windows\AuserGestioneTrasporti.exe"
& $exePath $csvFile $outputFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "Transformation failed!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Transformation completed!" -ForegroundColor Green
Write-Host ""

# Analyze the CSV to show what we expect
Write-Host "=== CSV Analysis - Destinazione Column ===" -ForegroundColor Cyan
Write-Host "Showing first 10 rows with their expected Destinazione values..." -ForegroundColor Yellow
Write-Host ""

$csv = Import-Csv $csvFile -Delimiter ';'
$count = 0

foreach ($row in $csv) {
    if ($count -ge 10) { break }
    if ([string]::IsNullOrWhiteSpace($row.'COMUNE DESTINAZIONE')) { continue }
    
    $count++
    
    $parts = @()
    if (![string]::IsNullOrWhiteSpace($row.'COMUNE DESTINAZIONE')) {
        $parts += $row.'COMUNE DESTINAZIONE'.Trim()
    }
    if (![string]::IsNullOrWhiteSpace($row.'INDIRIZZO DESTINAZIONE')) {
        $parts += $row.'INDIRIZZO DESTINAZIONE'.Trim()
    }
    if (![string]::IsNullOrWhiteSpace($row.'CAUSALE DESTINAZIONE')) {
        $parts += $row.'CAUSALE DESTINAZIONE'.Trim()
    }
    
    $expectedDestinazione = $parts -join ' '
    
    Write-Host "Row $count`: $($row.COGNOME) $($row.NOME)" -ForegroundColor White
    Write-Host "  COMUNE DESTINAZIONE: '$($row.'COMUNE DESTINAZIONE')'" -ForegroundColor Gray
    Write-Host "  INDIRIZZO DESTINAZIONE: '$($row.'INDIRIZZO DESTINAZIONE')'" -ForegroundColor Gray
    Write-Host "  CAUSALE DESTINAZIONE: '$($row.'CAUSALE DESTINAZIONE')'" -ForegroundColor Gray
    Write-Host "  Expected Destinazione: '$expectedDestinazione'" -ForegroundColor Green
    Write-Host ""
}

Write-Host ""
Write-Host "=== Verification Instructions ===" -ForegroundColor Cyan
Write-Host "1. Open the output file: $outputFile" -ForegroundColor White
Write-Host "2. Check column 5 (Destinazione)" -ForegroundColor White
Write-Host "3. Verify that each row contains: COMUNE + INDIRIZZO + CAUSALE DESTINAZIONE" -ForegroundColor White
Write-Host ""
Write-Host "Output file created: $outputFile" -ForegroundColor Green
