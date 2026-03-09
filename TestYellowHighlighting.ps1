# Test script to verify yellow highlighting functionality
# This script checks if the CSV file contains the expected rows with "Accompag. con macchina attrezzata"

$csvPath = "168514-Estrazione_1772702766454.csv"

Write-Host "Testing Yellow Highlighting Functionality" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Read CSV file
$csvContent = Import-Csv -Path $csvPath -Delimiter ";"

# Filter rows with "Accompag. con macchina attrezzata"
$highlightRows = $csvContent | Where-Object { $_.ATTIVITA -like "*Accompag. con macchina attrezzata*" }

Write-Host "Total rows in CSV: $($csvContent.Count)" -ForegroundColor Yellow
Write-Host "Rows with 'Accompag. con macchina attrezzata': $($highlightRows.Count)" -ForegroundColor Yellow
Write-Host ""

Write-Host "Details of rows that should be highlighted:" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

foreach ($row in $highlightRows) {
    $status = $row.'DESCRIZIONE STATO SERVIZIO'
    $date = $row.'DATA SERVIZIO'
    $time = $row.'ORA INIZIO SERVIZIO'
    $name = "$($row.'COGNOME ASSISTITO') $($row.'NOME ASSISTITO')"
    
    Write-Host ""
    Write-Host "Date: $date, Time: $time" -ForegroundColor White
    Write-Host "Patient: $name" -ForegroundColor White
    Write-Host "Status: $status" -ForegroundColor $(if ($status -eq "ANNULLATO") { "Red" } else { "Green" })
    
    if ($status -eq "ANNULLATO") {
        Write-Host "  -> This row will be FILTERED OUT (not highlighted)" -ForegroundColor Red
    } else {
        Write-Host "  -> This row WILL BE HIGHLIGHTED in yellow" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "========" -ForegroundColor Cyan

$annullato = ($highlightRows | Where-Object { $_.'DESCRIZIONE STATO SERVIZIO' -eq "ANNULLATO" }).Count
$toHighlight = $highlightRows.Count - $annullato

Write-Host "Total rows with 'Accompag. con macchina attrezzata': $($highlightRows.Count)" -ForegroundColor White
Write-Host "Rows with ANNULLATO status (filtered out): $annullato" -ForegroundColor Red
Write-Host "Rows that WILL BE HIGHLIGHTED: $toHighlight" -ForegroundColor Yellow
Write-Host ""
Write-Host "Expected behavior: $toHighlight rows should have yellow background in the Excel output" -ForegroundColor Green
