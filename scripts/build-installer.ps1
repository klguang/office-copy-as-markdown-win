param(
    [string]$Profile = "win-x64-framework-dependent"
)

$ErrorActionPreference = "Stop"

function Resolve-IsccPath {
    $command = Get-Command iscc -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
    ) | Where-Object { $_ -and (Test-Path $_) }

    $candidateList = @($candidates)
    if ($candidateList.Count -gt 0) {
        return $candidateList[0]
    }

    throw "Inno Setup 6 is required to build the installer. Install ISCC.exe or add it to PATH."
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$publishDir = Join-Path $repoRoot "artifacts\publish\$Profile"
$installerScript = Join-Path $repoRoot "packaging\installer\OfficeCopyAsMarkdown.iss"
$installerOutputDir = Join-Path $repoRoot "artifacts\installer"
$projectFile = Join-Path $repoRoot "src\OfficeCopyAsMarkdown\OfficeCopyAsMarkdown.csproj"

if (-not (Test-Path (Join-Path $publishDir "OfficeCopyAsMarkdown.exe"))) {
    throw "Published application was not found at $publishDir. Run publish.ps1 first."
}

if (-not (Test-Path $installerScript)) {
    throw "Installer script not found at $installerScript."
}

[xml]$projectXml = Get-Content $projectFile -Encoding UTF8
$version = $projectXml.Project.PropertyGroup.Version | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($version)) {
    throw "Unable to determine application version from $projectFile."
}

New-Item -ItemType Directory -Force -Path $installerOutputDir | Out-Null
$iscc = Resolve-IsccPath

& $iscc `
    "/DAppVersion=$version" `
    "/DSourceDir=$publishDir" `
    "/DOutputDir=$installerOutputDir" `
    $installerScript

Write-Host "Installer built in $installerOutputDir"
