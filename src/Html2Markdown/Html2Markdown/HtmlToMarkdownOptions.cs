namespace Html2Markdown;

public sealed class HtmlToMarkdownOptions
{
    public bool ExtractClipboardFragment { get; init; }

    public HtmlToMarkdownRepairMode RepairMode { get; init; } = HtmlToMarkdownRepairMode.IfSourceTextAvailable;

    public byte[]? FallbackImagePng { get; init; }

    public string? SourceText { get; init; }

    public IHtmlDialectAdapter? DialectAdapter { get; init; }

    public IHeadingInferenceStrategy? HeadingInferenceStrategy { get; init; }

    public Action<HtmlToMarkdownLogLevel, string>? Log { get; init; }
}
