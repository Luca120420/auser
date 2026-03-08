$excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx"

if (!(Test-Path $excelPath)) {
    Write-Host "File not found: $excelPath"
    exit
}

# Use .NET directly
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Excel files are ZIP archives
$zipPath = [System.IO.Path]::GetTempFileName()
Copy-Item $excelPath $zipPath -Force

try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    
    Write-Host "=== CHECKING FOR LABORATORI SHEET ==="
    Write-Host "File: $excelPath"
    Write-Host ""
    Write-Host "Worksheets in file:"
    
    $worksheets = $zip.Entries | Where-Object { $_.FullName -like "xl/worksheets/sheet*.xml" }
    
    # Read workbook.xml to get sheet names
    $workbookEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/workbook.xml" }
    if ($workbookEntry) {
        $stream = $workbookEntry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()
        
        # Extract sheet names using regex
        $matches = [regex]::Matches($content, 'name="([^"]+)"')
        foreach ($match in $matches) {
            $sheetName = $match.Groups[1].Value
            Write-Host "  - '$sheetName'"
            
            if ($sheetName -eq "laboratori") {
                Write-Host ""
                Write-Host "✓ FOUND 'laboratori' sheet!"
                $zip.Dispose()
                Remove-Item $zipPath -Force
                exit 0
            }
        }
    }
    
    Write-Host ""
    Write-Host "❌ NO 'laboratori' sheet found!"
    
    $zip.Dispose()
} finally {
    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
}
