# 用户指南

English version: [user-guide.md](user-guide.md)

`office-copy-as-markdown` 是一个 Windows 桌面小工具，用来把 Word 或 OneNote 桌面版中的当前选中内容复制为 Markdown，并把转换结果直接写回剪贴板。

## 项目定位

这个项目不是 Office 内嵌加载项，而是一个独立运行的桌面辅助程序。

工作方式：

1. 程序驻留在系统托盘。
2. 注册全局快捷键，默认是 `Ctrl+Shift+C`。
3. 只有当前台窗口是 Word 或 OneNote 时才触发。
4. 触发后向前台 Office 窗口发送复制操作。
5. 读取剪贴板中的 `HTML Format`、纯文本和图片数据。
6. 转换成 Markdown 后重新写回剪贴板。
7. 你可以直接在 Markdown 编辑器里粘贴。

## 支持范围

当前支持的常见格式包括：

- 标题
- 段落
- 粗体
- 斜体
- 删除线
- 引用
- 有序列表
- 无序列表
- 任务列表
- 行内代码
- 代码块
- 链接
- 图片
- 表格
- 分隔线

## 是否必须同时安装 Word 和 OneNote

不需要。

程序只会在前台进程是以下任意一个时生效：

- `WINWORD`
- `ONENOTE`

因此：

- 只安装 Word，可以正常使用
- 只安装 OneNote，可以正常使用
- 两者都安装，也可以正常使用
- 两者都没安装，程序只是驻留，不会触发转换

## 安装程序

推荐给最终用户直接运行安装包：

```text
.\artifacts\installer\OfficeCopyAsMarkdown-Setup-<version>.exe
```

安装程序特性：

- 仅安装到当前用户
- 写入“应用和功能”里的卸载项
- 程序安装到 `%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown`
- 卸载时默认保留 `%LOCALAPPDATA%\OfficeCopyAsMarkdown` 下的日志和设置
- 依赖 Microsoft Windows Desktop Runtime 10（x64），缺失时会在安装前提示

## 使用方法

1. 启动程序。
2. 打开 Word 或 OneNote 桌面版。
3. 选中内容。
4. 按当前配置的快捷键，默认是 `Ctrl+Shift+C`。
5. 在 Markdown 文档中直接粘贴。

你可以通过托盘菜单里的 `Settings...` 修改快捷键。

## 已知限制

- 本项目依赖 Word 和 OneNote 写入剪贴板的格式质量，Office 输出的 HTML 越完整，转换效果越好。
- Office 生成的 HTML 会因版本和内容复杂度不同而变化，极复杂排版仍可能需要手工微调。
- 当 HTML 中没有可直接复用的图片地址，但剪贴板里存在位图时，程序会退化为 `data:image/png;base64,...` 的 Markdown 图片。
- 这类未签名、会注册快捷键并读写剪贴板的桌面工具，可能触发杀软启发式检测；默认推荐使用非单文件版本。

## 相关文档

- [标题判断规则](heading-rules.zh-CN.md)
- [全文完整性校验与补回规则](content-check.zh-CN.md)
- [开发与发布](development.zh-CN.md)
- [故障排查与日志](troubleshooting.zh-CN.md)
