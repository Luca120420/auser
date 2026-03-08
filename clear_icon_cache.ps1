# Clear Windows Icon Cache
Write-Host "Clearing Windows icon cache..."

# Stop Explorer
Stop-Process -Name explorer -Force

# Wait a moment
Start-Sleep -Seconds 2

# Delete icon cache files
$iconCachePath = "$env:LOCALAPPDATA\IconCache.db"
$iconCacheFolder = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"

if (Test-Path $iconCachePath) {
    Remove-Item $iconCachePath -Force
    Write-Host "Deleted IconCache.db"
}

if (Test-Path $iconCacheFolder) {
    Get-ChildItem $iconCacheFolder -Filter "iconcache*" | Remove-Item -Force
    Get-ChildItem $iconCacheFolder -Filter "thumbcache*" | Remove-Item -Force
    Write-Host "Deleted icon cache files from Explorer folder"
}

# Restart Explorer
Start-Process explorer.exe

Write-Host "Icon cache cleared! Please check the executable icon now."
