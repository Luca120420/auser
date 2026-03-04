# Package AuserExcelTransformer into auser.zip with only runtime files

$sourceDir = "bin/Debug/net9.0-windows"
$zipFile = "auser.zip"

# Remove existing zip if it exists
if (Test-Path $zipFile) {
    Remove-Item $zipFile -Force
    Write-Host "Removed existing $zipFile"
}

# Create temporary directory for packaging
$tempDir = "temp_package"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Essential runtime files (exclude test-related DLLs)
$essentialFiles = @(
    "AuserExcelTransformer.exe",
    "AuserExcelTransformer.dll",
    "AuserExcelTransformer.runtimeconfig.json",
    "AuserExcelTransformer.deps.json",
    "EPPlus.dll",
    "EPPlus.Interfaces.dll",
    "EPPlus.System.Drawing.dll",
    "CsvHelper.dll",
    "Microsoft.Extensions.Configuration.dll",
    "Microsoft.Extensions.Configuration.Abstractions.dll",
    "Microsoft.Extensions.Configuration.FileExtensions.dll",
    "Microsoft.Extensions.Configuration.Json.dll",
    "Microsoft.Extensions.FileProviders.Abstractions.dll",
    "Microsoft.Extensions.FileProviders.Physical.dll",
    "Microsoft.Extensions.FileSystemGlobbing.dll",
    "Microsoft.Extensions.Primitives.dll",
    "Microsoft.IO.RecyclableMemoryStream.dll",
    "Newtonsoft.Json.dll"
)

# Copy essential files
foreach ($file in $essentialFiles) {
    $sourcePath = Join-Path $sourceDir $file
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath -Destination $tempDir
        Write-Host "Copied: $file"
    } else {
        Write-Host "Warning: $file not found" -ForegroundColor Yellow
    }
}

# Copy Italian localization (only the application resources)
$itDir = Join-Path $tempDir "it"
New-Item -ItemType Directory -Path $itDir | Out-Null
$itResourceFile = Join-Path $sourceDir "it/AuserExcelTransformer.resources.dll"
if (Test-Path $itResourceFile) {
    Copy-Item $itResourceFile -Destination $itDir
    Write-Host "Copied: it/AuserExcelTransformer.resources.dll"
}

# Create zip file
Write-Host "`nCreating $zipFile..."
Compress-Archive -Path "$tempDir/*" -DestinationPath $zipFile -Force

# Clean up temp directory
Remove-Item $tempDir -Recurse -Force

# Show result
if (Test-Path $zipFile) {
    $zipSize = (Get-Item $zipFile).Length / 1MB
    Write-Host "`nSuccess! Created $zipFile ($([math]::Round($zipSize, 2)) MB)" -ForegroundColor Green
    Write-Host "`nContents:"
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $zipFile))
    $zip.Entries | ForEach-Object { Write-Host "  $($_.FullName)" }
    $zip.Dispose()
} else {
    Write-Host "Error: Failed to create $zipFile" -ForegroundColor Red
}
