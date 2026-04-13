# Office Copy as Markdown

Chinese version: [README.md](../README.md)

Copy the current Word or OneNote selection as Markdown with one shortcut.

## Supported Office Clients

Supported clients:

- Word desktop
- OneNote desktop

The app only runs when the foreground window is `WINWORD` or `ONENOTE`, and it does not require both to be installed.

## Installer

Installer file name:

```text
OfficeCopyAsMarkdown-Setup-<version>.exe
```

Installer behavior:

- installs for the current user only
- writes an Apps & Features uninstall entry
- installs binaries to `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- preserves user data under `%LOCALAPPDATA%\OfficeCopyAsMarkdown` on uninstall
- requires Microsoft Windows Desktop Runtime 10 (x64), and prompts before install if it is missing

## Usage

1. Start `OfficeCopyAsMarkdown.exe`.
2. Open Word or OneNote desktop.
3. Select content.
4. Press the hotkey. The default is `Ctrl+Shift+C`.
5. Paste into a Markdown document.

You can change it from `Settings...` in the tray menu.

## App Notes

- This is a standalone desktop app, not an Office web add-in.
- The app stays in the system tray and registers a global hotkey.
- The converted result is written back to the clipboard.

## Supported Markdown Output

Supported Markdown output:

- headings, paragraphs, blockquotes
- bold, italic, strikethrough
- ordered lists, unordered lists, task lists
- fenced code blocks, links, images, tables, horizontal rules

## Notes And Limitations

- Conversion quality depends on the `HTML Format` data emitted by Office.
- Images may fall back to `data:image/png;base64,...` when no reusable source is available.
- Complex layouts may still need manual cleanup.
- Unsigned desktop apps can trigger antivirus heuristics.

## Related Documents

- Chinese guide: [README.md](../README.md)
- [Heading rules](heading-rules.md)
- [Content completeness and backfill rules](content-check.md)
- [Development and publishing](development.md)
- [Troubleshooting and logging](troubleshooting.md)
