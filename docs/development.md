# Development And Publishing

Chinese version: [development.zh-CN.md](development.zh-CN.md)

This document collects build, publish, and development-oriented installation details that are intentionally kept out of the root `README.md`.

## Prerequisites

- Windows
- .NET SDK 10 or later

## Build

```powershell
dotnet build .\OfficeCopyAsMarkdown.slnx
```

## Publish

Publish the recommended profile:

```powershell
.\scripts\publish.ps1
```

That command builds both:

- the framework-dependent published app in `.\artifacts\publish\win-x64-framework-dependent`
- a per-user Inno Setup installer in `.\artifacts\installer`

If Inno Setup 6 is not installed, either install `ISCC.exe` first or skip the installer step:

```powershell
.\scripts\publish.ps1 -SkipInstaller
```

Other publish modes:

```powershell
.\scripts\publish.ps1 -SelfContained
.\scripts\publish.ps1 -SelfContained -SingleFile
```

Build just the installer from an existing published output:

```powershell
.\scripts\build-installer.ps1
```

## Development Installation

For development or manual deployment, install the published app with the helper script:

```powershell
.\scripts\install.ps1
```

Optional: create a Startup shortcut so it launches when you sign in:

```powershell
.\scripts\install.ps1 -Startup
```

The script copies files to:

```text
%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown
```

If you want to install a self-contained profile explicitly:

```powershell
.\scripts\install.ps1 -Profile win-x64-self-contained
.\scripts\install.ps1 -Profile win-x64-self-contained-single-file
```

## Packaging Notes

- The default recommendation is the framework-dependent profile.
- A self-contained single-file build is larger and may be more likely to trigger antivirus heuristics.

## Microsoft References

- [OneNote add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/)
- [Custom keyboard shortcuts in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- [HTML Clipboard Format](https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format)
