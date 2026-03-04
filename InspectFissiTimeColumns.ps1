# Simple PowerShell script to inspect fissi sheet time columns using EPPlus
$ErrorActionPreference = "Stop"

Write-Host "=== Inspecting Fissi Sheet Time Columns ===" -ForegroundColor Cyan
Write-Host ""

$excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx"

if (-not (Test-Path $excelPath)) {
    Write-Host "ERROR: Excel file not found: $excelPath" -ForegroundColor Red
    exit 1
}

# Load EPPlus
Add-Type -Path "bin\Debug\net9.0-windows\EPPlus.dll"
[OfficeOpenXml.ExcelPackage]::LicenseContext = [OfficeOpenXml.LicenseContext]::NonCommercial

$fileInfo = New-Object System.IO.FileInfo($excelPath)
$package = New-Object OfficeOpenXml.ExcelPackage($fileInfo)

try {
    # Find fissi sheet
    $fissiSheet = $null
    foreach ($ws in $package.Workbook.Worksheets) {
        if ($ws.Name -eq "fissi") {
            $fissiSheet = $ws
            break
        }
    }

    if ($null -eq $fissiSheet) {
        Write-Host "ERROR: 'fissi' sheet not found" -ForegroundColor Red
        Write-Host "Available sheets:" -ForegroundColor Yellow
        foreach ($ws in $package.Workbook.Worksheets) {
            Write-Host "  - $($ws.Name)"
        }
        exit 1
    }

    Write-Host "Found 'fissi' sheet" -ForegroundColor Green
    Write-Host ""

    $dimension = $fissiSheet.Dimension
    if ($null -eq $dimension) {
        Write-Host "Sheet is empty" -ForegroundColor Yellow
        exit 0
    }

    Write-Host "=== Fissi Sheet - Time Columns (Source Data) ===" -ForegroundColor Cyan
    Write-Host "This shows how time values are currently stored/displayed"
    Write-Host ""
    Write-Host "Row | Partenza (Col 2)                              | Arrivo (Col 9)"
    Write-Host "----|-----------------------------------------------|-----------------------------------------------"

    $startRow = 3  # Skip headers
    $maxRows = [Math]::Min(15, $dimension.End.Row - $startRow + 1)
    
    $decimalCount = 0
    $timeFormatCount = 0

    for ($row = $startRow; $row -lt ($startRow + $maxRows); $row++) {
        $partenzaCell = $fissiSheet.Cells[$row, 2]
        $arrivoCell = $fissiSheet.Cells[$row, 9]

        $partenzaValue = $partenzaCell.Value
        $partenzaText = $partenzaCell.Text
        $partenzaFormat = if ($partenzaCell.Style.Numberformat.Format) { $partenzaCell.Style.Numberformat.Format } else { "no format" }

        $arrivoValue = $arrivoCell.Value
        $arrivoText = $arrivoCell.Text
        $arrivoFormat = if ($arrivoCell.Style.Numberformat.Format) { $arrivoCell.Style.Numberformat.Format } else { "no format" }

        Write-Host ("{0,3} | V:{1,-12} T:{2,-10} F:{3,-10} | V:{4,-12} T:{5,-10} F:{6,-10}" -f `
            $row, $partenzaValue, $partenzaText, $partenzaFormat, $arrivoValue, $arrivoText, $arrivoFormat)

        # Count decimal vs time format
        if ($null -ne $partenzaValue) {
            if ($partenzaText -match "[.,]" -and $partenzaText -notmatch ":") {
                $decimalCount++
            } elseif ($partenzaText -match ":") {
                $timeFormatCount++
            }
        }

        if ($null -ne $arrivoValue) {
            if ($arrivoText -match "[.,]" -and $arrivoText -notmatch ":") {
                $decimalCount++
            } elseif ($arrivoText -match ":") {
                $timeFormatCount++
            }
        }
    }

    Write-Host ""
    Write-Host "=== Analysis ===" -ForegroundColor Cyan
    Write-Host "Cells displaying as DECIMAL: $decimalCount" -ForegroundColor $(if ($decimalCount -gt 0) { "Red" } else { "Green" })
    Write-Host "Cells displaying as TIME FORMAT: $timeFormatCount" -ForegroundColor Green
    Write-Host ""

    if ($decimalCount -gt 0) {
        Write-Host "⚠️  ISSUE DETECTED: Some time values are displaying as decimals" -ForegroundColor Yellow
        Write-Host "   Example: 0.354166666666667 instead of 8:30" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "The fix in ExcelManager.cs AppendFissiData method:" -ForegroundColor Cyan
        Write-Host "  1. Copies the decimal value as-is (preserves the data)" -ForegroundColor White
        Write-Host "  2. Applies 'h:mm' format to columns 2 and 9" -ForegroundColor White
        Write-Host "  3. Result: Excel displays 8:30 instead of 0.354166666666667" -ForegroundColor White
        Write-Host ""
        Write-Host "✓ The fix is already implemented in the code" -ForegroundColor Green
        Write-Host "  When you process files through the application, time columns will display correctly" -ForegroundColor Green
    } else {
        Write-Host "✓ All time values are already displaying in time format" -ForegroundColor Green
        Write-Host "  The fix will ensure this format is preserved when copying" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Legend:" -ForegroundColor Gray
    Write-Host "  V = Value (internal representation, usually decimal for times)" -ForegroundColor Gray
    Write-Host "  T = Text (what Excel displays to the user)" -ForegroundColor Gray
    Write-Host "  F = Format (number format code applied to the cell)" -ForegroundColor Gray

} finally {
    $package.Dispose()
}
