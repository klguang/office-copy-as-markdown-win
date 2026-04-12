$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "Programs\OfficeCopyAsMarkdown"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Office Copy as Markdown"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Office Copy as Markdown.lnk"
$startupShortcut = Join-Path ([Environment]::GetFolderPath("Startup")) "Office Copy as Markdown.lnk"

if (Test-Path $installDir) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
}

if (Test-Path $startMenuDir) {
    Remove-Item -LiteralPath $startMenuDir -Recurse -Force
}

if (Test-Path $desktopShortcut) {
    Remove-Item -LiteralPath $desktopShortcut -Force
}

if (Test-Path $startupShortcut) {
    Remove-Item -LiteralPath $startupShortcut -Force
}

Write-Host "Office Copy as Markdown has been removed."
