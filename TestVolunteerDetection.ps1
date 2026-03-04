# Test script to verify volunteer detection fix
Write-Host "Testing volunteer detection with settimana_nuova_lombardi.xlsx" -ForegroundColor Cyan

# Run the application in diagnostic mode
$exePath = "bin\Release\net9.0-windows\AuserExcelTransformer.exe"

if (Test-Path $exePath) {
    Write-Host "Application found at: $exePath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Please test the volunteer email feature with the following:" -ForegroundColor Yellow
    Write-Host "  1. Excel file: settimana_nuova_lombardi.xlsx" -ForegroundColor White
    Write-Host "  2. Sheet: 6" -ForegroundColor White
    Write-Host "  3. Expected volunteer: Lombardi" -ForegroundColor White
    Write-Host "  4. Expected email: luca.lombardi120@gmail.com" -ForegroundColor White
    Write-Host ""
    Write-Host "The fix should now:" -ForegroundColor Yellow
    Write-Host "  - Detect that row 2 contains the 'Volontario' column header" -ForegroundColor White
    Write-Host "  - Read volunteer data starting from row 3" -ForegroundColor White
    Write-Host "  - Find 'Lombardi' in the Volontario column" -ForegroundColor White
    Write-Host "  - Send an email to luca.lombardi120@gmail.com" -ForegroundColor White
    Write-Host ""
    
    # Start the application
    Start-Process $exePath
} else {
    Write-Host "Error: Application not found at $exePath" -ForegroundColor Red
    Write-Host "Please build the application first with: dotnet build --configuration Release" -ForegroundColor Yellow
}
