# Test script to verify italic formatting for laboratori sheet columns
# This script runs the application and checks if Assistito and Indirizzo columns are formatted in italic

Write-Host "Testing laboratori sheet italic formatting..." -ForegroundColor Cyan
Write-Host ""

# Check if the executable exists
$exePath = "bin\Release\net9.0-windows\AuserExcelTransformer.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "ERROR: Executable not found at $exePath" -ForegroundColor Red
    Write-Host "Please build the project first with: dotnet build --configuration Release" -ForegroundColor Yellow
    exit 1
}

# Check if test file exists
$testFile = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx"
if (-not (Test-Path $testFile)) {
    Write-Host "ERROR: Test file not found: $testFile" -ForegroundColor Red
    exit 1
}

Write-Host "Test file found: $testFile" -ForegroundColor Green
Write-Host ""
Write-Host "MANUAL TEST INSTRUCTIONS:" -ForegroundColor Yellow
Write-Host "1. Run the application: $exePath" -ForegroundColor White
Write-Host "2. Open the test file: $testFile" -ForegroundColor White
Write-Host "3. Select a sheet that has laboratori data (look for sheets with 'laboratori' in the name)" -ForegroundColor White
Write-Host "4. Process the file and create a new sheet" -ForegroundColor White
Write-Host "5. In the output sheet, verify that:" -ForegroundColor White
Write-Host "   - Column 3 (Assistito) has italic text for rows from laboratori sheet" -ForegroundColor Cyan
Write-Host "   - Column 4 (Indirizzo) has italic text for rows from laboratori sheet" -ForegroundColor Cyan
Write-Host "   - Rows from fissi sheet should NOT have italic formatting" -ForegroundColor Cyan
Write-Host ""
Write-Host "The italic formatting implementation is in Services/ExcelManager.cs" -ForegroundColor Gray
Write-Host "Lines 688-691 in the AppendLaboratoriData method" -ForegroundColor Gray
Write-Host ""
Write-Host "Press any key to launch the application..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Launch the application
Start-Process $exePath
