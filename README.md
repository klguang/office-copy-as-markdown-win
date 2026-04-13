# Office Copy as Markdown

English version: [user-guide.md](docs/user-guide.md)

把 Word / OneNote 选中内容一键复制为 Markdown。

## 快速了解

- 支持 Word 桌面版和 OneNote 桌面版
- 默认快捷键是 `Ctrl+Shift+C`
- 推荐从 [Releases](https://github.com/klguang/office-copy-as-markdown-win/releases) 下载安装包

## 安装

安装包文件名：

```text
OfficeCopyAsMarkdown-Setup-<version>.exe
```

安装程序特性：

- 仅安装到当前用户
- 程序安装到 `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- 卸载时保留 `%LOCALAPPDATA%\OfficeCopyAsMarkdown` 下的用户数据
- 依赖 Microsoft Windows Desktop Runtime 10（x64）

## 使用方法

1. 启动 `OfficeCopyAsMarkdown.exe`。
2. 打开 Word 或 OneNote 桌面版。
3. 选中内容。
4. 按快捷键，默认是 `Ctrl+Shift+C`。
5. 在 Markdown 文档中粘贴。

可通过托盘图标中的 `Settings...` 修改快捷键。

## 程序说明

- 这是一个独立桌面程序，不是 Office Web Add-in。
- 支持 Word / OneNote 桌面版，仅在前台窗口是 `WINWORD` / `ONENOTE` 时触发。
- 程序常驻系统托盘，触发后会把结果写回剪贴板。

## 支持的 Markdown 输出：

支持标题、段落、列表、引用、强调、代码、链接、图片和表格等常见 Markdown 内容。

## 说明与限制

- 转换质量依赖 Office 提供的 `HTML Format` 数据。
- 图片无法复用原始来源时，会回退为 `data:image/png;base64,...`。
- 复杂排版或杀毒软件误报属于已知限制。

## 相关文档

- English guide: [docs/user-guide.md](docs/user-guide.md)
- 标题规则: [EN](docs/heading-rules.md) / [中文](docs/heading-rules.zh-CN.md)
- 开发与发布: [EN](docs/development.md) / [中文](docs/development.zh-CN.md)
- 故障排查: [EN](docs/troubleshooting.md) / [中文](docs/troubleshooting.zh-CN.md)
