param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [switch]$SelfContained,
    [switch]$SingleFile,
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repoRoot "src\OfficeCopyAsMarkdown\OfficeCopyAsMarkdown.csproj"
$profileName = if ($SelfContained) {
    if ($SingleFile) { "$Runtime-self-contained-single-file" } else { "$Runtime-self-contained" }
} else {
    "$Runtime-framework-dependent"
}

$output = Join-Path $repoRoot "artifacts\publish\$profileName"

dotnet publish $project `
    -c $Configuration `
    -r $Runtime `
    --self-contained:$($SelfContained.IsPresent.ToString().ToLowerInvariant()) `
    /p:PublishSingleFile=$($SingleFile.IsPresent.ToString().ToLowerInvariant()) `
    /p:IncludeNativeLibrariesForSelfExtract=$($SingleFile.IsPresent.ToString().ToLowerInvariant()) `
    -o $output

Write-Host "Published to $output"

if (-not $SelfContained -and -not $SingleFile -and -not $SkipInstaller) {
    & (Join-Path $PSScriptRoot "build-installer.ps1") -Profile $profileName
}
