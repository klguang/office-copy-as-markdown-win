namespace Html2Markdown;

public static class HtmlToMarkdownPipeline
{
    public static string NormalizeHtml(string html) =>
        html ?? string.Empty;

    public static string ExtractFragment(string html) =>
        CfHtmlExtractor.ExtractFragment(html);

    public static string ConvertFragment(
        string html,
        byte[]? fallbackImagePng = null,
        Action<HtmlToMarkdownLogLevel, string>? log = null)
    {
        return ConvertFragment(
            html,
            new HtmlToMarkdownOptions
            {
                FallbackImagePng = fallbackImagePng,
                Log = log
            });
    }

    public static string ConvertFragment(string html, HtmlToMarkdownOptions? options)
    {
        options ??= new HtmlToMarkdownOptions();

        return HtmlToMarkdownLog.RunWith(options.Log, () =>
        {
            HtmlStructureLogger.LogFragmentStructure(html);
            return MarkdownConverter.Convert(
                html,
                options.FallbackImagePng,
                options.DialectAdapter ?? StandardHtmlDialectAdapter.Instance,
                options.HeadingInferenceStrategy ?? DefaultHeadingInferenceStrategy.Instance);
        });
    }

    public static HtmlToMarkdownRepairResult RepairMarkdown(
        string markdown,
        string? sourceText,
        Action<HtmlToMarkdownLogLevel, string>? log = null)
    {
        return HtmlToMarkdownLog.RunWith(log, () =>
        {
            var repaired = MarkdownContentGuard.RepairMarkdown(markdown, sourceText);
            return new HtmlToMarkdownRepairResult(repaired.IsComplete, repaired.Markdown, repaired.MissingLines);
        });
    }

    public static bool ShouldKeepMarkdown(
        string markdown,
        string? sourceText,
        Action<HtmlToMarkdownLogLevel, string>? log = null)
    {
        return HtmlToMarkdownLog.RunWith(log, () => MarkdownContentGuard.ShouldKeepMarkdown(markdown, sourceText));
    }

    public static string BuildConservativeMarkdown(string sourceText) =>
        MarkdownContentGuard.BuildConservativeMarkdown(sourceText);
}
