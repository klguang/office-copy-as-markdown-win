# 故障排查与日志

English version: [troubleshooting.md](troubleshooting.md)

## 日志

日志采用混合方案：

- `Debug` 构建默认开启详细日志。
- `Release` 构建默认不写日志。
- `Release` 版本可以通过环境变量 `OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL` 临时开启日志。

日志文件写入：

```text
%LOCALAPPDATA%\OfficeCopyAsMarkdown\logs
```

临时开启 `Release` 日志示例：

```powershell
$env:OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL = "Warning"
C:\Users\kevin\AppData\Local\Programs\OfficeCopyAsMarkdown\OfficeCopyAsMarkdown.exe
Remove-Item Env:OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL
```

可用值：

- `None`
- `Error`
- `Warning`
- `Information`
- `Debug`
- `Trace`

## 排查说明

- 转换质量取决于 Word 和 OneNote 写入剪贴板的格式。
- 当 Office 提供可用的 `HTML Format` 数据时，结果通常最好。
- 如果 Office 生成的 HTML 因版本或区域设置不同而变化，复杂排版可能需要在粘贴后手工微调。
- 如果安全软件对程序报毒，优先使用默认的 framework-dependent 发布配置，而不是体积更大的 self-contained 单文件版本。
