# office-copy-as-markdown

`office-copy-as-markdown` is a Windows helper for Word and OneNote desktop that converts the current selection to Markdown and puts the Markdown back on the clipboard.

中文说明见 [README.zh-CN.md](D:/workspace/office-copy-as-markdown/README.zh-CN.md)。
Heading rules are documented in [HEADING_RULES.md](D:/workspace/office-copy-as-markdown/HEADING_RULES.md).

## Why this is a desktop helper instead of an Office web add-in

Microsoft's current Office add-in documentation creates two constraints:

1. The OneNote JavaScript add-in model targets OneNote on the web, not the Win32 desktop client.
2. Office add-in keyboard shortcuts are documented for Office Add-ins and currently supported in Word and Excel, but that route does not provide a single desktop implementation that covers both Word and OneNote desktop in the same way.

Because of that, this project takes a host-agnostic desktop route:

1. Register `Ctrl+Shift+C`.
2. Only act when the foreground process is `WINWORD` or `ONENOTE`.
3. Send a normal `Ctrl+C` to the foreground Office window.
4. Read the clipboard's `HTML Format`, plain text, and bitmap data.
5. Convert the copied fragment to Markdown.
6. Replace the clipboard text with Markdown so you can paste directly into a Markdown editor.

## Supported Markdown output

The converter handles the common formats requested for Word and OneNote selections:

- headings
- paragraphs
- bold
- italic
- strikethrough
- blockquotes
- ordered and unordered lists
- task lists
- inline code
- fenced code blocks
- links
- images
- tables
- horizontal rules

## Build

Prerequisites:

- Windows
- .NET SDK 10 or later

Build:

```powershell
dotnet build .\OfficeCopyAsMarkdown.slnx
```

Publish the recommended profile:

```powershell
.\scripts\publish.ps1
```

The published output is written to:

```text
.\artifacts\publish\win-x64-framework-dependent
```

Other publish modes:

```powershell
.\scripts\publish.ps1 -SelfContained
.\scripts\publish.ps1 -SelfContained -SingleFile
```

## Install

Install the published app for the current user:

```powershell
.\scripts\install.ps1
```

Optional: create a Startup shortcut so it launches when you sign in:

```powershell
.\scripts\install.ps1 -Startup
```

The installer copies files to:

```text
%LOCALAPPDATA%\OfficeCopyAsMarkdown
```

If you want to install a self-contained profile explicitly:

```powershell
.\scripts\install.ps1 -Profile win-x64-self-contained
.\scripts\install.ps1 -Profile win-x64-self-contained-single-file
```

## Usage

1. Start `OfficeCopyAsMarkdown.exe`.
2. Open Word or OneNote desktop.
3. Select content.
4. Press `Ctrl+Shift+C`.
5. Paste into a Markdown document.

The hotkey is only honored when Word or OneNote is the foreground window.

## Logging

Mixed logging is used:

- `Debug` builds log verbosely by default.
- `Release` builds do not log by default.
- `Release` logging can be enabled temporarily with the `OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL` environment variable.

Log files are written to:

```text
%LOCALAPPDATA%\OfficeCopyAsMarkdown\logs
```

Example for a temporary `Release` session:

```powershell
$env:OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL = "Warning"
.\OfficeCopyAsMarkdown.exe
Remove-Item Env:OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL
```

Accepted values:

- `None`
- `Error`
- `Warning`
- `Information`
- `Debug`
- `Trace`

## Notes and limitations

- This project depends on the clipboard formats emitted by Word and OneNote. Results are best when Office provides rich `HTML Format` clipboard data.
- When clipboard HTML does not contain a usable image source but the clipboard includes a bitmap, the app embeds that image as a `data:image/png;base64,...` Markdown image so it can still be pasted directly.
- Office-generated HTML varies across versions and locales. The converter includes heuristics for Office list paragraphs and style-based formatting, but very complex layouts may still need manual cleanup.
- Unsigned desktop utilities that register hotkeys, simulate `Ctrl+C`, and rewrite clipboard contents may trigger antivirus heuristics. The default publish profile is unpacked and framework-dependent because that shape is usually less suspicious than a large self-contained single-file executable.

## Microsoft documentation referenced

- [OneNote add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/)
- [Custom keyboard shortcuts in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- [HTML Clipboard Format](https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format)
