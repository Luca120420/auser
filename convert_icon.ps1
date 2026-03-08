# Convert PNG to ICO using .NET
Add-Type -AssemblyName System.Drawing

$pngPath = "Resources\app_icon.png"
$icoPath = "Resources\app_icon.ico"

# Load the PNG image
$img = [System.Drawing.Image]::FromFile((Resolve-Path $pngPath))

# Create icon from image
$icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]$img).GetHicon())

# Save as ICO
$fileStream = [System.IO.File]::Create($icoPath)
$icon.Save($fileStream)
$fileStream.Close()

Write-Host "Icon converted successfully: $icoPath"
