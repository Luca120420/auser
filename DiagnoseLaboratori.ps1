# Load EPPlus assembly
Add-Type -Path "C:\Users\lucal\.nuget\packages\epplus\7.5.2\lib\net8.0\EPPlus.dll"

# Set license context
[OfficeOpenXml.ExcelPackage]::LicenseContext = [OfficeOpenXml.LicenseContext]::NonCommercial

# Find the most recent Excel file
$excelFiles = Get-ChildItem -Path . -Filter "*.xlsx" | 
    Where-Object { $_.Name -notlike "*~$*" } |
    Sort-Object LastWriteTime -Descending

if ($excelFiles.Count -eq 0) {
    Write-Host "No Excel files found in current directory"
    exit
}

$latestFile = $excelFiles[0]
Write-Host "Inspecting: $($latestFile.Name)"
Write-Host "Last modified: $($latestFile.LastWriteTime)"
Write-Host ""

$package = New-Object OfficeOpenXml.ExcelPackage($latestFile.FullName)

# Check if laboratori sheet exists
$laboratoriSheet = $package.Workbook.Worksheets | Where-Object { $_.Name -eq "laboratori" }

if ($null -eq $laboratoriSheet) {
    Write-Host "No 'laboratori' sheet found in workbook"
    $sheetNames = $package.Workbook.Worksheets | ForEach-Object { $_.Name }
    Write-Host "Available sheets: $($sheetNames -join ', ')"
    $package.Dispose()
    exit
}

Write-Host "Found 'laboratori' sheet"
Write-Host ""

# Check dimensions
$dimension = $laboratoriSheet.Dimension
if ($null -eq $dimension) {
    Write-Host "Laboratori sheet is empty (no dimension)"
    $package.Dispose()
    exit
}

Write-Host "Sheet dimensions: $($dimension.Address)"
Write-Host "Rows: $($dimension.Start.Row) to $($dimension.End.Row)"
Write-Host "Columns: $($dimension.Start.Column) to $($dimension.End.Column)"
Write-Host ""

# Check row 1
Write-Host "Row 1 (useless header):"
for ($col = 1; $col -le [Math]::Min(10, $dimension.End.Column); $col++) {
    $value = $laboratoriSheet.Cells[1, $col].Text
    Write-Host "  Col $col - '$value'"
}
Write-Host ""

# Check row 2 (should be column headers)
Write-Host "Row 2 (column headers):"
for ($col = 1; $col -le [Math]::Min(10, $dimension.End.Column); $col++) {
    $value = $laboratoriSheet.Cells[2, $col].Text
    Write-Host "  Col $col - '$value'"
}
Write-Host ""

# Check if row 2, col 1 contains "Data"
$row2Col1 = $laboratoriSheet.Cells[2, 1].Text
Write-Host "Row 2, Col 1 value: '$row2Col1'"
Write-Host "Equals 'Data' (case-insensitive): $($row2Col1 -eq 'Data')"
Write-Host ""

# Count data rows (starting from row 3)
$dataRowCount = 0
$emptyDataRows = 0

Write-Host "Data rows (starting from row 3):"
for ($row = 3; $row -le $dimension.End.Row; $row++) {
    $dataValue = $laboratoriSheet.Cells[$row, 1].Text
    
    if ([string]::IsNullOrWhiteSpace($dataValue)) {
        $emptyDataRows++
        Write-Host "  Row $row - EMPTY (would be skipped)"
    }
    else {
        $dataRowCount++
        $partenza = $laboratoriSheet.Cells[$row, 2].Text
        $assistito = $laboratoriSheet.Cells[$row, 3].Text
        $avv = $laboratoriSheet.Cells[$row, 10].Text
        Write-Host "  Row $row - Data='$dataValue', Partenza='$partenza', Assistito='$assistito', Avv='$avv'"
    }
}

Write-Host ""
Write-Host "Summary:"
Write-Host "  Total data rows with non-empty Data column: $dataRowCount"
Write-Host "  Rows with empty Data column (skipped): $emptyDataRows"
Write-Host "  Expected rows in output: $dataRowCount"

# Check the output sheet
Write-Host ""
$outputSheet = $package.Workbook.Worksheets | Where-Object { 
    $_.Name -like "*nuovo*" -or $_.Name -like "*output*" -or $_.Name -like "*new*" 
} | Select-Object -First 1

if ($null -ne $outputSheet) {
    Write-Host "Found output sheet: '$($outputSheet.Name)'"
    $outputDim = $outputSheet.Dimension
    if ($null -ne $outputDim) {
        Write-Host "  Output rows: $($outputDim.Start.Row) to $($outputDim.End.Row)"
        
        # Check for laboratori data in output (look for Avv column values)
        $laboratoriRowsFound = 0
        for ($row = 3; $row -le $outputDim.End.Row; $row++) {
            $avvValue = $outputSheet.Cells[$row, 10].Text
            if (-not [string]::IsNullOrWhiteSpace($avvValue)) {
                $laboratoriRowsFound++
            }
        }
        Write-Host "  Rows with Avv column data (laboratori indicator): $laboratoriRowsFound"
    }
}
else {
    Write-Host "No output sheet found"
}

$package.Dispose()
