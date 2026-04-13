# Troubleshooting And Logging

Chinese version: [troubleshooting.zh-CN.md](troubleshooting.zh-CN.md)

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
C:\Users\kevin\AppData\Local\Programs\OfficeCopyAsMarkdown\OfficeCopyAsMarkdown.exe
Remove-Item Env:OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL
```

Accepted values:

- `None`
- `Error`
- `Warning`
- `Information`
- `Debug`
- `Trace`

## Troubleshooting Notes

- Conversion quality depends on the clipboard formats emitted by Word and OneNote.
- Results are usually best when Office provides usable `HTML Format` data.
- When Office-generated HTML varies across versions or locales, complex layouts may need manual cleanup after paste.
- If security software flags the app, prefer the default framework-dependent publish profile rather than a large self-contained single-file build.
