# office-copy-as-markdown

`office-copy-as-markdown` is a Windows helper for Word and OneNote desktop that converts the current selection to Markdown and puts the Markdown back on the clipboard.

## What It Supports

Supported Office clients:

- Word desktop
- OneNote desktop

Supported Markdown output:

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

## Install

Recommended for end users: run the generated installer:

```text
.\artifacts\installer\OfficeCopyAsMarkdown-Setup-<version>.exe
```

The installer:

- installs for the current user only
- writes an Apps & Features uninstall entry
- installs binaries to `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- preserves `%LOCALAPPDATA%\OfficeCopyAsMarkdown` user data on uninstall
- requires Microsoft Windows Desktop Runtime 10 (x64)

## Usage

1. Start `OfficeCopyAsMarkdown.exe`.
2. Open Word or OneNote desktop.
3. Select content.
4. Press the hotkey. The default is `Ctrl+Shift+C`.
5. Paste into a Markdown document.

The hotkey is only honored when Word or OneNote is the foreground window.
You can change it from the tray icon through `Settings...`.

## Notes And Limitations

- Results are best when Office provides rich `HTML Format` clipboard data.
- If the clipboard only contains a bitmap for an image, the app falls back to a `data:image/png;base64,...` Markdown image.
- Very complex Office layouts may still need manual cleanup after conversion.
- Unsigned desktop utilities that register hotkeys and rewrite clipboard contents may trigger antivirus heuristics.

## Documentation

- User guide: [English](docs/user-guide.md) / [zh-CN](docs/user-guide.zh-CN.md)
- Heading rules: [English](docs/heading-rules.md) / [zh-CN](docs/heading-rules.zh-CN.md)
- Content completeness rules: [English](docs/content-check.md) / [zh-CN](docs/content-check.zh-CN.md)
- Development and publishing: [English](docs/development.md) / [zh-CN](docs/development.zh-CN.md)
- Troubleshooting and logging: [English](docs/troubleshooting.md) / [zh-CN](docs/troubleshooting.zh-CN.md)
