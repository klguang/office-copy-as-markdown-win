# User Guide

Chinese version: [user-guide.zh-CN.md](user-guide.zh-CN.md)

`office-copy-as-markdown` is a Windows desktop helper that copies the current selection from Word or OneNote desktop, converts it to Markdown, and writes the converted result back to the clipboard.

## Positioning

This project is not an Office web add-in. It is a standalone desktop helper.

How it works:

1. The app stays in the system tray.
2. It registers a global hotkey. The default is `Ctrl+Shift+C`.
3. It only triggers when the foreground window is Word or OneNote.
4. It sends a normal copy command to the foreground Office window.
5. It reads `HTML Format`, plain text, and image data from the clipboard.
6. It converts the copied content to Markdown and writes it back to the clipboard.
7. You can paste the result directly into a Markdown editor.

## Supported Output

Common supported formats include:

- headings
- paragraphs
- bold
- italic
- strikethrough
- blockquotes
- ordered lists
- unordered lists
- task lists
- inline code
- fenced code blocks
- links
- images
- tables
- horizontal rules

## Do You Need Both Word And OneNote

No.

The app only runs when the foreground process is one of these:

- `WINWORD`
- `ONENOTE`

That means:

- if only Word is installed, it still works
- if only OneNote is installed, it still works
- if both are installed, it still works
- if neither is installed, the app only stays resident and never triggers

## Installer

For end users, the recommended option is to run the generated installer:

```text
.\artifacts\installer\OfficeCopyAsMarkdown-Setup-<version>.exe
```

Installer behavior:

- installs for the current user only
- writes an Apps & Features uninstall entry
- installs binaries to `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- preserves logs and settings under `%LOCALAPPDATA%\OfficeCopyAsMarkdown` on uninstall
- requires Microsoft Windows Desktop Runtime 10 (x64), and prompts before install if it is missing

## Usage

1. Start the app.
2. Open Word or OneNote desktop.
3. Select content.
4. Press the configured hotkey. The default is `Ctrl+Shift+C`.
5. Paste into a Markdown document.

You can change the hotkey from `Settings...` in the tray menu.

## Known Limitations

- Conversion quality depends on the clipboard data emitted by Word and OneNote. Results are best when Office provides rich HTML output.
- Office-generated HTML varies by version and content complexity, so very complex layouts may still need manual cleanup.
- When the clipboard HTML does not contain a reusable image source but the clipboard includes a bitmap, the app falls back to a `data:image/png;base64,...` Markdown image.
- Unsigned desktop tools that register hotkeys and read or rewrite clipboard contents can trigger antivirus heuristics. The default recommendation is the non-single-file build.

## Related Documents

- [Heading rules](heading-rules.md)
- [Content completeness and backfill rules](content-check.md)
- [Development and publishing](development.md)
- [Troubleshooting and logging](troubleshooting.md)
