param(
    [switch]$Startup,
    [string]$Profile = "win-x64-framework-dependent"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$publishDir = Join-Path $repoRoot "artifacts\publish\$Profile"
$exePath = Join-Path $publishDir "OfficeCopyAsMarkdown.exe"

if (-not (Test-Path $exePath)) {
    if ($Profile -like "*self-contained*") {
        $singleFile = $Profile -like "*single-file*"
        & (Join-Path $PSScriptRoot "publish.ps1") -SelfContained -SingleFile:$singleFile
    } else {
        & (Join-Path $PSScriptRoot "publish.ps1") -SkipInstaller
    }
}

$installDir = Join-Path $env:LOCALAPPDATA "Programs\OfficeCopyAsMarkdown"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item -Path (Join-Path $publishDir "*") -Destination $installDir -Recurse -Force

$installedExe = Join-Path $installDir "OfficeCopyAsMarkdown.exe"
$shell = New-Object -ComObject WScript.Shell

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Office Copy as Markdown"
New-Item -ItemType Directory -Force -Path $startMenuDir | Out-Null
$startMenuShortcut = Join-Path $startMenuDir "Office Copy as Markdown.lnk"
$startMenu = $shell.CreateShortcut($startMenuShortcut)
$startMenu.TargetPath = $installedExe
$startMenu.WorkingDirectory = $installDir
$startMenu.Save()

$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Office Copy as Markdown.lnk"
$desktop = $shell.CreateShortcut($desktopShortcut)
$desktop.TargetPath = $installedExe
$desktop.WorkingDirectory = $installDir
$desktop.Save()

if ($Startup) {
    $startupShortcut = Join-Path ([Environment]::GetFolderPath("Startup")) "Office Copy as Markdown.lnk"
    $startup = $shell.CreateShortcut($startupShortcut)
    $startup.TargetPath = $installedExe
    $startup.WorkingDirectory = $installDir
    $startup.Save()
}

Write-Host "Installed to $installDir"
