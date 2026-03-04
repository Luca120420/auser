# PowerShell script to test time format fix with real files
# This script uses the compiled application DLL to process the files

$ErrorActionPreference = "Stop"

Write-Host "=== Testing Time Format Fix with Real Files ===" -ForegroundColor Cyan
Write-Host ""

# File paths
$csvPath = "168514-Estrazione_1770193162042.csv"
$excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx"
$outputPath = "test_output_time_format_verification.xlsx"
$dllPath = "bin\Debug\net9.0-windows\AuserExcelTransformer.dll"

# Verify files exist
if (-not (Test-Path $csvPath)) {
    Write-Host "ERROR: CSV file not found: $csvPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $excelPath)) {
    Write-Host "ERROR: Excel file not found: $excelPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: DLL not found: $dllPath" -ForegroundColor Red
    Write-Host "Please run 'dotnet build' first" -ForegroundColor Yellow
    exit 1
}

Write-Host "Input files verified:" -ForegroundColor Green
Write-Host "  CSV: $csvPath"
Write-Host "  Excel: $excelPath"
Write-Host "  Output: $outputPath"
Write-Host ""

# Load the assembly
Write-Host "Loading application assembly..." -ForegroundColor Yellow
Add-Type -Path $dllPath
Add-Type -Path "bin\Debug\net9.0-windows\EPPlus.dll"

# Set EPPlus license
[OfficeOpenXml.ExcelPackage]::LicenseContext = [OfficeOpenXml.LicenseContext]::NonCommercial

Write-Host "Assembly loaded successfully" -ForegroundColor Green
Write-Host ""

try {
    # Create service instances
    Write-Host "Step 1: Initializing services..." -ForegroundColor Yellow
    $csvParser = New-Object AuserExcelTransformer.Services.CSVParser
    $rulesEngine = New-Object AuserExcelTransformer.Services.TransformationRulesEngine
    $dataTransformer = New-Object AuserExcelTransformer.Services.DataTransformer($rulesEngine)
    $excelManager = New-Object AuserExcelTransformer.Services.ExcelManager
    Write-Host "  Services initialized" -ForegroundColor Green

    # Parse CSV
    Write-Host "Step 2: Parsing CSV..." -ForegroundColor Yellow
    $appointments = $csvParser.ParseCSV($csvPath)
    Write-Host "  Parsed $($appointments.Count) appointments" -ForegroundColor Green

    # Transform data
    Write-Host "Step 3: Transforming data..." -ForegroundColor Yellow
    $transformationResult = $dataTransformer.Transform($appointments)
    Write-Host "  Transformed $($transformationResult.Rows.Count) rows" -ForegroundColor Green

    # Load Excel workbook
    Write-Host "Step 4: Loading Excel workbook..." -ForegroundColor Yellow
    $workbook = $excelManager.LoadExcelFile($excelPath)
    Write-Host "  Loaded workbook with $($workbook.Sheets.Count) sheets" -ForegroundColor Green

    # Find fissi sheet
    Write-Host "Step 5: Finding 'fissi' sheet..." -ForegroundColor Yellow
    $fissiSheet = $null
    foreach ($sheet in $workbook.Sheets) {
        if ($sheet.Name -eq "fissi") {
            $fissiSheet = $sheet
            break
        }
    }
    
    if ($null -eq $fissiSheet) {
        Write-Host "  ERROR: 'fissi' sheet not found" -ForegroundColor Red
        exit 1
    }
    Write-Host "  Found 'fissi' sheet" -ForegroundColor Green

    # Create new sheet
    Write-Host "Step 6: Creating new sheet..." -ForegroundColor Yellow
    $newSheet = $excelManager.CreateNewSheet($workbook, $transformationResult.HeaderInfo)
    Write-Host "  Created new sheet" -ForegroundColor Green

    # Append CSV data
    Write-Host "Step 7: Appending CSV data..." -ForegroundColor Yellow
    $excelManager.AppendTransformedData($newSheet, $transformationResult)
    Write-Host "  Appended CSV data" -ForegroundColor Green

    # Append fissi data
    Write-Host "Step 8: Appending fissi data..." -ForegroundColor Yellow
    $startRow = $transformationResult.Rows.Count + 2
    $excelManager.AppendFissiData($newSheet, $fissiSheet, $startRow)
    Write-Host "  Appended fissi data starting at row $startRow" -ForegroundColor Green

    # Save workbook
    Write-Host "Step 9: Saving workbook..." -ForegroundColor Yellow
    $excelManager.SaveExcelFile($workbook, $outputPath)
    Write-Host "  Saved to $outputPath" -ForegroundColor Green

    Write-Host ""
    Write-Host "=== SUCCESS ===" -ForegroundColor Green
    Write-Host "Output file created: $outputPath"
    Write-Host ""
    Write-Host "Now verifying time format in output..." -ForegroundColor Cyan
    Write-Host ""

    # Verify time format in output
    $outputWorkbook = $excelManager.LoadExcelFile($outputPath)
    $outputSheet = $outputWorkbook.Sheets[0]
    $worksheet = $outputSheet.Worksheet

    Write-Host "=== Time Column Verification ===" -ForegroundColor Cyan
    Write-Host "Checking rows starting from $startRow"
    Write-Host ""
    Write-Host "Row | Partenza (Col 2)                              | Arrivo (Col 9)"
    Write-Host "----|-----------------------------------------------|-----------------------------------------------"

    $rowsChecked = 0
    $rowsWithCorrectFormat = 0
    $maxRows = [Math]::Min(10, $worksheet.Dimension.End.Row - $startRow + 1)

    for ($row = $startRow; $row -lt ($startRow + $maxRows); $row++) {
        $partenzaCell = $worksheet.Cells[$row, 2]
        $arrivoCell = $worksheet.Cells[$row, 9]

        $partenzaValue = $partenzaCell.Value
        $partenzaText = $partenzaCell.Text
        $partenzaFormat = $partenzaCell.Style.Numberformat.Format

        $arrivoValue = $arrivoCell.Value
        $arrivoText = $arrivoCell.Text
        $arrivoFormat = $arrivoCell.Style.Numberformat.Format

        Write-Host ("{0,3} | V:{1,-12} T:{2,-10} F:{3,-10} | V:{4,-12} T:{5,-10} F:{6,-10}" -f `
            $row, $partenzaValue, $partenzaText, $partenzaFormat, $arrivoValue, $arrivoText, $arrivoFormat)

        if ($null -ne $partenzaValue -or $null -ne $arrivoValue) {
            $rowsChecked++
            
            if ($partenzaFormat -eq "h:mm" -or $arrivoFormat -eq "h:mm") {
                $rowsWithCorrectFormat++
            }
        }
    }

    Write-Host ""
    Write-Host "=== Results ===" -ForegroundColor Cyan
    Write-Host "Rows checked: $rowsChecked"
    Write-Host "Rows with correct time format: $rowsWithCorrectFormat"
    Write-Host ""
    
    if ($rowsWithCorrectFormat -eq $rowsChecked -and $rowsChecked -gt 0) {
        Write-Host "✓ SUCCESS: All time columns have correct 'h:mm' format!" -ForegroundColor Green
        Write-Host "✓ Time values display as time (e.g., 8:30) instead of decimal (e.g., 0.354166666666667)" -ForegroundColor Green
    } else {
        Write-Host "⚠ WARNING: Some rows may not have correct format" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Legend: V=Value (internal), T=Text (displayed), F=Format" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Output file location: $(Resolve-Path $outputPath)" -ForegroundColor Cyan

} catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    exit 1
}
