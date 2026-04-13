# Office Copy as Markdown

English version: [user-guide.md](docs/user-guide.md)

把 Word / OneNote 选中内容一键复制为 Markdown。

## 支持的 Office 客户端

支持的客户端：

- Word 桌面版
- OneNote 桌面版

程序仅在前台窗口是 `WINWORD` 或 `ONENOTE` 时触发，不需要同时安装两者。

## 安装

安装包文件名：

```text
OfficeCopyAsMarkdown-Setup-<version>.exe
```

安装程序特性：

- 仅安装到当前用户
- 写入“应用和功能”里的卸载项
- 程序安装到 `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- 卸载时保留 `%LOCALAPPDATA%\OfficeCopyAsMarkdown` 下的用户数据
- 依赖 Microsoft Windows Desktop Runtime 10（x64），缺失时安装程序会先提示

## 使用方法

1. 启动 `OfficeCopyAsMarkdown.exe`。
2. 打开 Word 或 OneNote 桌面版。
3. 选中内容。
4. 按快捷键，默认是 `Ctrl+Shift+C`。
5. 在 Markdown 文档中粘贴。

可通过托盘图标中的 `Settings...` 修改快捷键。

## 程序说明

- 这是一个独立桌面程序，不是 Office Web Add-in。
- 程序常驻系统托盘，并注册全局快捷键。
- 转换结果会直接写回剪贴板。

## 支持的 Markdown 输出：

- 标题、段落、引用
- 粗体、斜体、删除线
- 有序列表、无序列表、任务列表
- 围栏代码块、链接、图片、表格、分隔线

## 说明与限制

- 转换质量依赖 Office 提供的 `HTML Format` 数据。
- 图片无法复用原始来源时，会回退为 `data:image/png;base64,...`。
- 复杂排版转换后仍可能需要手工调整。
- 未签名桌面程序可能触发杀毒软件的启发式检测。

## 相关文档

- English guide: [docs/user-guide.md](docs/user-guide.md)
- 标题规则: [EN](docs/heading-rules.md) / [中文](docs/heading-rules.zh-CN.md)
- 内容完整性: [EN](docs/content-check.md) / [中文](docs/content-check.zh-CN.md)
- 开发与发布: [EN](docs/development.md) / [中文](docs/development.zh-CN.md)
- 故障排查: [EN](docs/troubleshooting.md) / [中文](docs/troubleshooting.zh-CN.md)
