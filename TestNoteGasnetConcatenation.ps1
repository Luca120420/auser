# Test script to verify Note Gasnet concatenation functionality
# This tests that DESCRIZIONE PUNTO PARTENZA is concatenated to NOTE E RICHIESTE

Write-Host "Testing Note Gasnet Concatenation..." -ForegroundColor Cyan
Write-Host ""

# Build the project
Write-Host "Building project..." -ForegroundColor Yellow
dotnet build AuserExcelTransformer.csproj -c Debug

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed!" -ForegroundColor Red
    exit 1
}

Write-Host "Build successful!" -ForegroundColor Green
Write-Host ""

# Run the transformation
$csvFile = "168514-Estrazione_1772702766454.csv"
$outputFile = "test_output_note_gasnet.xlsx"

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
Write-Host "=== CSV Analysis ===" -ForegroundColor Cyan
Write-Host "Looking for rows with DESCRIZIONE PUNTO PARTENZA values..." -ForegroundColor Yellow
Write-Host ""

$csv = Import-Csv $csvFile -Delimiter ';'
$rowsWithDescPartenza = $csv | Where-Object { 
    $_.'DESCRIZIONE PUNTO PARTENZA' -and 
    $_.'DESCRIZIONE PUNTO PARTENZA' -ne 'null' -and 
    $_.'DESCRIZIONE PUNTO PARTENZA'.Trim() -ne ''
}

if ($rowsWithDescPartenza.Count -eq 0) {
    Write-Host "No rows found with DESCRIZIONE PUNTO PARTENZA values (all are 'null' or empty)" -ForegroundColor Yellow
} else {
    Write-Host "Found $($rowsWithDescPartenza.Count) row(s) with DESCRIZIONE PUNTO PARTENZA:" -ForegroundColor Green
    Write-Host ""
    
    foreach ($row in $rowsWithDescPartenza) {
        Write-Host "Row: $($row.COGNOME) $($row.NOME)" -ForegroundColor White
        Write-Host "  DESCRIZIONE PUNTO PARTENZA: '$($row.'DESCRIZIONE PUNTO PARTENZA')'" -ForegroundColor Cyan
        Write-Host "  NOTE E RICHIESTE: '$($row.'NOTE E RICHIESTE')'" -ForegroundColor Cyan
        Write-Host "  Expected Note Gasnet: '$($row.'NOTE E RICHIESTE') $($row.'DESCRIZIONE PUNTO PARTENZA')'" -ForegroundColor Green
        Write-Host ""
    }
}

Write-Host ""
Write-Host "=== Verification Instructions ===" -ForegroundColor Cyan
Write-Host "1. Open the output file: $outputFile" -ForegroundColor White
Write-Host "2. Check column 12 (Note Gasnet) for the rows listed above" -ForegroundColor White
Write-Host "3. Verify that DESCRIZIONE PUNTO PARTENZA is concatenated to NOTE E RICHIESTE" -ForegroundColor White
Write-Host "4. For rows with 'null' or empty DESCRIZIONE PUNTO PARTENZA, only NOTE E RICHIESTE should appear" -ForegroundColor White
Write-Host ""
Write-Host "Output file created: $outputFile" -ForegroundColor Green
