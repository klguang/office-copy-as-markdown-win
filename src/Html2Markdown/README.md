# Html2Markdown

`Html2Markdown` 是一个面向通用 HTML 输入的 Markdown 转换库。

它的职责很明确：

- 接收 HTML 或 CF_HTML 片段
- 转换为可读的 Markdown
- 在需要时结合源文本做保守补全和内容完整性校验
- 提供可扩展的 HTML 方言适配点和标题推断策略

它默认只理解标准 HTML 语义，不内置 Office 专属规则。像 Word、OneNote 这类来源的特殊 HTML 结构，如果需要兼容，应由调用方通过 `IHtmlDialectAdapter` 在应用层补充。

## 设计目标

- 保持库本身与具体宿主解耦
- 默认行为对“普通 HTML”可直接使用
- 对复杂来源保留扩展能力，而不是把来源方言写死在库里
- 标题推断默认可用，但允许调用方完全覆盖
- 当 HTML 结果不完整时，允许回退到更保守的 Markdown

## 不做什么

- 不直接依赖 Office、Interop、剪贴板或 WinForms
- 不默认识别 `mso-*`、`HeadingN`、`MsoListParagraph` 之类 Office 特征
- 不负责上层应用的 HTML 预清洗策略
- 不承诺把任意脏 HTML 修复为完美 Markdown

## 包结构

当前包目录下的关键入口：

- `HtmlToMarkdownConverter`
  对外主入口。适合直接从 HTML 得到最终 Markdown。
- `HtmlToMarkdownPipeline`
  对外工具型入口。适合按步骤使用：片段提取、片段转换、完整性校验、保守补全。
- `HtmlToMarkdownOptions`
  转换参数对象。
- `IHtmlDialectAdapter`
  非标准 HTML 方言扩展点。
- `IHeadingInferenceStrategy`
  标题推断扩展点。
- `StandardHtmlDialectAdapter`
  默认 HTML 方言适配器。
- `DefaultHeadingInferenceStrategy`
  默认标题推断实现。

## 快速开始

### 1. 直接转换普通 HTML

```csharp
using Html2Markdown;

var html = """
<html>
<body>
  <h2>Title</h2>
  <p>Hello <strong>world</strong>.</p>
</body>
</html>
""";

var markdown = HtmlToMarkdownConverter.Convert(html);
```

输出大致为：

```md
## Title

Hello **world**.
```

### 2. 转换 CF_HTML

如果输入是剪贴板常见的 `CF_HTML` 格式，可以开启片段提取：

```csharp
var markdown = HtmlToMarkdownConverter.Convert(
    cfHtml,
    new HtmlToMarkdownOptions
    {
        ExtractClipboardFragment = true
    });
```

### 3. 启用源文本补全

当 HTML 转换结果可能丢字、丢行、丢列表项时，可以提供源文本：

```csharp
var markdown = HtmlToMarkdownConverter.Convert(
    html,
    new HtmlToMarkdownOptions
    {
        SourceText = sourceText,
        RepairMode = HtmlToMarkdownRepairMode.IfSourceTextAvailable
    });
```

这会在 Markdown 初稿生成后做一次完整性检查。如果检查发现内容缺失，库会先尝试修补；修补仍不完整时，会回退到更保守的 Markdown。

## 公开 API

### `HtmlToMarkdownConverter`

```csharp
string Convert(string html, HtmlToMarkdownOptions? options = null)
```

这是面向调用方的主入口，适合“一次调用拿最终结果”的场景。

行为概览：

- 输入为空时，如果提供了 `FallbackImagePng`，会生成内嵌 `data:image/png;base64,...` 的 Markdown 图片
- 可选提取 CF_HTML 的 `StartFragment/EndFragment`
- 执行 HTML 到 Markdown 的主转换
- 可选执行基于 `SourceText` 的修补和回退

### `HtmlToMarkdownPipeline`

适合更细粒度的场景。

#### `NormalizeHtml`

```csharp
string NormalizeHtml(string html)
```

当前实现是轻量透传，主要作为统一入口保留。

#### `ExtractFragment`

```csharp
string ExtractFragment(string html)
```

用于从 CF_HTML 中提取片段。

#### `ConvertFragment`

```csharp
string ConvertFragment(string html, HtmlToMarkdownOptions? options)
string ConvertFragment(
    string html,
    byte[]? fallbackImagePng = null,
    Action<HtmlToMarkdownLogLevel, string>? log = null)
```

用于把一个已知片段直接转换成 Markdown。这个入口不会再做 `ExtractClipboardFragment` 的额外判断，通常适合你已经知道输入就是目标片段的场景。

#### `RepairMarkdown`

```csharp
HtmlToMarkdownRepairResult RepairMarkdown(
    string markdown,
    string? sourceText,
    Action<HtmlToMarkdownLogLevel, string>? log = null)
```

用于检查 Markdown 是否覆盖了源文本中的可见内容，并在必要时尝试补回缺失行。

#### `ShouldKeepMarkdown`

```csharp
bool ShouldKeepMarkdown(
    string markdown,
    string? sourceText,
    Action<HtmlToMarkdownLogLevel, string>? log = null)
```

用于只做完整性判断，不做修补。

#### `BuildConservativeMarkdown`

```csharp
string BuildConservativeMarkdown(string sourceText)
```

直接从源文本构造“尽量不丢内容”的保守 Markdown，不做标题推断。

## `HtmlToMarkdownOptions`

```csharp
public sealed class HtmlToMarkdownOptions
{
    public bool ExtractClipboardFragment { get; init; }
    public HtmlToMarkdownRepairMode RepairMode { get; init; }
    public byte[]? FallbackImagePng { get; init; }
    public string? SourceText { get; init; }
    public IHtmlDialectAdapter? DialectAdapter { get; init; }
    public IHeadingInferenceStrategy? HeadingInferenceStrategy { get; init; }
    public Action<HtmlToMarkdownLogLevel, string>? Log { get; init; }
}
```

字段说明：

- `ExtractClipboardFragment`
  为 `true` 时，先按 CF_HTML 提取片段，再进入转换。
- `RepairMode`
  当前支持：
  - `None`
  - `IfSourceTextAvailable`
- `FallbackImagePng`
  当 HTML 没有可用文本，但上层有 PNG 二进制时，可回退成 Markdown 图片。
- `SourceText`
  原始纯文本，用于完整性检查和保守回退。
- `DialectAdapter`
  自定义 HTML 方言适配器。未提供时使用 `StandardHtmlDialectAdapter.Instance`。
- `HeadingInferenceStrategy`
  自定义标题推断策略。未提供时使用 `DefaultHeadingInferenceStrategy.Instance`。
- `Log`
  可选日志回调。

## 默认 HTML 语义边界

`Html2Markdown` 默认只识别标准 HTML 结构：

- 语义标题：`<h1>` 到 `<h6>`
- 列表：`<ul>`、`<ol>`、`<li>`
- 引用：`<blockquote>`
- 代码块：`<pre>` 或具备明显预格式化/等宽特征的块
- 内联样式：链接、图片、粗体、斜体、删除线、行内代码、复选框等

默认不识别的来源特征包括但不限于：

- `mso-style-name: Heading N`
- `HeadingN`
- `mso-list`
- `MsoListParagraph`
- Office 条件注释和 Office XML 命名空间标签

这类来源特征如果需要支持，应交给调用方自己的预处理和方言适配器。

## 默认标题推断规则

库内置了 `DefaultHeadingInferenceStrategy`，它在没有语义标题时，尝试通过字号和样式推断标题层级。

### 语义标题优先

如果片段中已经存在语义标题：

- 标准 HTML 的 `<h1>` 到 `<h6>`
- 或方言适配器通过 `TryGetSemanticHeadingLevel(...)` 标出的标题

则整个片段禁用候选标题推断。

### 候选标题筛选

默认实现会筛掉以下内容：

- 结构容器
- 多视觉行内容
- 被列表、表格、引用、代码块包裹的内容
- 被当前方言适配器标记为“阻止标题推断”的内容
- 空文本
- 超过 `30` 个字符的文本
- 以 `。 . ， , ； ;` 结尾的文本
- 无法提取字号的文本

### 正文字号基线

正文基线字号取“当前片段中最常见的字号”。

### 有效候选条件

候选标题要成为有效候选，需满足以下任一条件：

- 字号至少比正文大 `2pt`
- 字号至少比正文大 `15%`
- 整段为粗体，且字号不小于正文

### 级别映射

默认参数：

- 最大支持级别数：`4`
- 稀疏映射起始级别：`2`

含义是：

- 最多只按字号分出 4 档
- 如果实际字号档位不足 4 档，则从 `##` 开始映射，而不是强行占用 `#`

这套规则是默认实现，不是硬约束。你可以通过自定义 `IHeadingInferenceStrategy` 完全替换。

## 扩展点

### `IHtmlDialectAdapter`

```csharp
public interface IHtmlDialectAdapter
{
    bool TryGetSemanticHeadingLevel(HtmlNode node, out int level);
    bool IsQuoteBlock(HtmlNode node);
    bool TryConvertListLikeBlock(HtmlNode node, int quoteDepth, out string markdown);
    bool ShouldBlockHeadingInference(HtmlNode node);
}
```

适用场景：

- 你的 HTML 不是标准 HTML，而是某种来源方言
- 你希望补充“非标准语义标题”识别
- 你希望识别自定义引用块
- 你希望把某些来源特有结构当成列表
- 你希望阻止某些特殊节点进入标题推断

默认实现 `StandardHtmlDialectAdapter` 什么都不额外做。

### `IHeadingInferenceStrategy`

```csharp
public interface IHeadingInferenceStrategy
{
    IReadOnlyDictionary<HtmlNode, int> InferHeadingLevels(
        HtmlNode root,
        IHtmlDialectAdapter dialectAdapter);
}
```

适用场景：

- 你不想使用默认字号推断
- 你想引入自己的文本密度、样式优先级、文档上下文规则
- 你需要对特定来源定义另一套标题规则

返回值是“节点到标题级别”的映射，例如：

- `1` 表示 `#`
- `2` 表示 `##`
- 以此类推

## 扩展示例

### 自定义语义标题识别

下面示例把 `data-heading-level` 当成语义标题：

```csharp
using Html2Markdown;
using HtmlAgilityPack;

public sealed class DataHeadingDialectAdapter : IHtmlDialectAdapter
{
    public bool TryGetSemanticHeadingLevel(HtmlNode node, out int level)
    {
        return int.TryParse(node.GetAttributeValue("data-heading-level", string.Empty), out level);
    }

    public bool IsQuoteBlock(HtmlNode node) => false;

    public bool TryConvertListLikeBlock(HtmlNode node, int quoteDepth, out string markdown)
    {
        markdown = string.Empty;
        return false;
    }

    public bool ShouldBlockHeadingInference(HtmlNode node) => false;
}
```

调用方式：

```csharp
var markdown = HtmlToMarkdownConverter.Convert(
    html,
    new HtmlToMarkdownOptions
    {
        DialectAdapter = new DataHeadingDialectAdapter()
    });
```

### 自定义标题推断

下面示例把第一个 `<p>` 强制视为一级标题：

```csharp
using Html2Markdown;
using HtmlAgilityPack;

public sealed class ForceFirstParagraphHeadingStrategy : IHeadingInferenceStrategy
{
    public IReadOnlyDictionary<HtmlNode, int> InferHeadingLevels(
        HtmlNode root,
        IHtmlDialectAdapter dialectAdapter)
    {
        var firstParagraph = root.Descendants("p").First();
        return new Dictionary<HtmlNode, int>
        {
            [firstParagraph] = 1
        };
    }
}
```

调用方式：

```csharp
var markdown = HtmlToMarkdownConverter.Convert(
    html,
    new HtmlToMarkdownOptions
    {
        HeadingInferenceStrategy = new ForceFirstParagraphHeadingStrategy()
    });
```

## 与 Office 相关的推荐用法

`Html2Markdown` 默认不处理 Office 方言。

如果你的上层输入来自 Word、OneNote 或其他 Office 导出的 HTML，推荐做法是：

1. 在应用层先做来源清洗
2. 在应用层提供自己的 `IHtmlDialectAdapter`
3. 把适配器通过 `HtmlToMarkdownOptions` 传入

换句话说：

- Office 兼容是“宿主责任”
- 通用 HTML 到 Markdown 是“库责任”

这样可以保持库本身足够纯净，也便于将来独立拆仓或发布 NuGet 包。

## 日志

如果你想观察转换过程，可以传入 `Log`：

```csharp
var markdown = HtmlToMarkdownConverter.Convert(
    html,
    new HtmlToMarkdownOptions
    {
        Log = (level, message) => Console.WriteLine($"[{level}] {message}")
    });
```

当前日志主要用于：

- HTML 结构观察
- 标题推断过程
- 完整性检查和回退决策

## 已知限制

- 默认标题推断依赖字号样本，片段过短时效果可能不稳定
- 保守补全优先保证“不丢内容”，不保证 Markdown 结构最优
- 对非常脏或非常碎片化的 HTML，输出仍可能需要上层二次处理
- 默认实现没有做来源识别，任何来源特性都不会自动启用

## 适合的使用方式

推荐按下面思路使用：

- 普通网页 HTML：直接 `HtmlToMarkdownConverter.Convert(...)`
- 剪贴板 CF_HTML：`ExtractClipboardFragment = true`
- 强调内容完整性：提供 `SourceText`
- 来自特殊来源：自定义 `IHtmlDialectAdapter`
- 有自己的标题规则：自定义 `IHeadingInferenceStrategy`

## 总结

`Html2Markdown` 的核心原则是：

- 默认只做标准 HTML
- 默认提供一套可用的标题推断
- 允许调用方覆盖来源方言和标题策略
- 不把宿主业务规则写死进库里

如果你希望把它作为独立项目拆出，这份边界设计就是为了服务那个目标。
