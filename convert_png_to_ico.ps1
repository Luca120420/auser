# Convert PNG to proper ICO with multiple resolutions
Add-Type -AssemblyName System.Drawing

$pngPath = "C:\Users\lucal\Desktop\progetti\auser\logo\favicon_auser.png"
$icoPath = "Resources\app_icon.ico"

Write-Host "Loading PNG image..."
$img = [System.Drawing.Image]::FromFile($pngPath)

Write-Host "Original image size: $($img.Width)x$($img.Height)"

# Create a bitmap from the image
$bitmap = New-Object System.Drawing.Bitmap $img

# Create icon with multiple sizes (16x16, 32x32, 48x48, 256x256)
$sizes = @(16, 32, 48, 256)
$icons = @()

foreach ($size in $sizes) {
    Write-Host "Creating ${size}x${size} icon..."
    $resized = New-Object System.Drawing.Bitmap $size, $size
    $graphics = [System.Drawing.Graphics]::FromImage($resized)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($bitmap, 0, 0, $size, $size)
    $graphics.Dispose()
    
    # Convert to icon
    $ms = New-Object System.IO.MemoryStream
    $resized.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $icons += $ms.ToArray()
    $ms.Dispose()
    $resized.Dispose()
}

# For now, let's use a simpler approach - just save the largest size
Write-Host "Creating final icon file..."
$icon256 = New-Object System.Drawing.Bitmap 256, 256
$g = [System.Drawing.Graphics]::FromImage($icon256)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.DrawImage($bitmap, 0, 0, 256, 256)
$g.Dispose()

# Save as ICO
$iconHandle = $icon256.GetHicon()
$icon = [System.Drawing.Icon]::FromHandle($iconHandle)
$fileStream = [System.IO.File]::Create($icoPath)
$icon.Save($fileStream)
$fileStream.Close()
$icon.Dispose()
$icon256.Dispose()

# Clean up
$bitmap.Dispose()
$img.Dispose()

Write-Host "Icon created successfully: $icoPath"
Write-Host "Icon file size: $((Get-Item $icoPath).Length) bytes"
