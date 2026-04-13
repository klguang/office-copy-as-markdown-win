# Office Copy as Markdown

Chinese version: [README.md](../README.md)

Copy the current Word or OneNote selection as Markdown with one shortcut.

## Quick Overview

- Supports Word desktop and OneNote desktop
- Default shortcut: `Ctrl+Shift+C`
- Recommended install source: [Releases](https://github.com/klguang/office-copy-as-markdown-win/releases)

## Installer

Installer file name:

```text
OfficeCopyAsMarkdown-Setup-<version>.exe
```

Installer behavior:

- installs for the current user only
- installs binaries to `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- preserves user data under `%LOCALAPPDATA%\OfficeCopyAsMarkdown` on uninstall
- requires Microsoft Windows Desktop Runtime 10 (x64)

## Usage

1. Start `OfficeCopyAsMarkdown.exe`.
2. Open Word or OneNote desktop.
3. Select content.
4. Press the hotkey. The default is `Ctrl+Shift+C`.
5. Paste into a Markdown document.

You can change it from `Settings...` in the tray menu.

## App Notes

- This is a standalone desktop app, not an Office web add-in.
- It supports Word / OneNote desktop and only runs when the foreground window is `WINWORD` / `ONENOTE`.
- The app stays in the system tray and writes the converted result back to the clipboard.

## Supported Markdown Output

Supports common Markdown content such as headings, paragraphs, lists, blockquotes, emphasis, code, links, images, and tables.

## Notes And Limitations

- Conversion quality depends on the `HTML Format` data emitted by Office.
- Images may fall back to `data:image/png;base64,...` when no reusable source is available.
- Complex layouts or antivirus false positives are known limitations.

## Related Documents

- Chinese guide: [README.md](../README.md)
- [Heading rules](heading-rules.md)
- [Development and publishing](development.md)
- [Troubleshooting and logging](troubleshooting.md)
