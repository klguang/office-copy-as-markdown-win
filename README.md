# 用户指南

English version: [user-guide.md](docs/user-guide.md)

`office-copy-as-markdown` 是一个 Windows 桌面小工具，用来把 Word 或 OneNote 桌面版中的当前选中内容复制为 Markdown，并把转换结果直接写回剪贴板。

## 支持范围

支持的 Office 客户端：

- Word 桌面版
- OneNote 桌面版

支持的 Markdown 输出：

- 标题
- 段落
- 粗体
- 斜体
- 删除线
- 引用
- 有序列表和无序列表
- 任务列表
- 行内代码
- 围栏代码块
- 链接
- 图片
- 表格
- 分隔线

## 安装

推荐普通用户从 [Releases](https://github.com/klguang/office-copy-as-markdown-win/releases) 下载最新安装包。

安装包文件名：

```text
OfficeCopyAsMarkdown-Setup-<version>.exe
```

安装程序特性：

- 仅安装到当前用户
- 写入“应用和功能”里的卸载项
- 程序安装到 `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- 卸载时保留 `%LOCALAPPDATA%\OfficeCopyAsMarkdown` 下的用户数据
- 依赖 Microsoft Windows Desktop Runtime 10（x64）

## 使用方法

1. 启动 `OfficeCopyAsMarkdown.exe`。
2. 打开 Word 或 OneNote 桌面版。
3. 选中内容。
4. 按快捷键，默认是 `Ctrl+Shift+C`。
5. 在 Markdown 文档中粘贴。

只有当前前台窗口是 Word 或 OneNote 时，快捷键才会生效。
你可以通过托盘图标中的 `Settings...` 修改快捷键。

## 说明与限制

- 当 Office 提供较完整的 `HTML Format` 剪贴板数据时，转换结果通常最好。
- 如果图片只有位图数据，程序会回退为 `data:image/png;base64,...` 形式的 Markdown 图片。
- 非常复杂的 Office 排版在转换后仍可能需要手工调整。
- 这类会注册快捷键并改写剪贴板的未签名桌面工具，可能触发杀毒软件的启发式检测。

## 相关文档

- 用户指南: [English](user-guide.md) / [zh-CN](user-guide.zh-CN.md)
- 标题规则: [English](heading-rules.md) / [zh-CN](heading-rules.zh-CN.md)
- 内容完整性规则: [English](content-check.md) / [zh-CN](content-check.zh-CN.md)
- 开发与发布: [English](development.md) / [zh-CN](development.zh-CN.md)
- 故障排查与日志: [English](troubleshooting.md) / [zh-CN](troubleshooting.zh-CN.md)