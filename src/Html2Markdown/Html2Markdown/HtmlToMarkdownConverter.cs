namespace Html2Markdown;

public static class HtmlToMarkdownConverter
{
    public static string Convert(string html, HtmlToMarkdownOptions? options = null)
    {
        options ??= new HtmlToMarkdownOptions();

        return HtmlToMarkdownLog.RunWith(options.Log, () =>
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return TryBuildImageOnlyMarkdown(options.FallbackImagePng);
            }

            var fragment = options.ExtractClipboardFragment
                ? CfHtmlExtractor.ExtractFragment(html)
                : html;

            HtmlStructureLogger.LogFragmentStructure(fragment);

            var markdown = MarkdownConverter.Convert(
                fragment,
                options.FallbackImagePng,
                options.DialectAdapter ?? StandardHtmlDialectAdapter.Instance,
                options.HeadingInferenceStrategy ?? DefaultHeadingInferenceStrategy.Instance);
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return FallbackToConservativeMarkdown(options, markdown);
            }

            if (options.RepairMode == HtmlToMarkdownRepairMode.None || string.IsNullOrWhiteSpace(options.SourceText))
            {
                return markdown;
            }

            var repaired = MarkdownContentGuard.RepairMarkdown(markdown, options.SourceText);
            if (repaired.IsComplete)
            {
                return repaired.Markdown;
            }

            HtmlToMarkdownLog.Warning("Markdown output remained incomplete after repair. Falling back to conservative Markdown.");
            return FallbackToConservativeMarkdown(options, markdown);
        });
    }

    private static string FallbackToConservativeMarkdown(HtmlToMarkdownOptions options, string currentMarkdown)
    {
        if (!string.IsNullOrWhiteSpace(options.SourceText))
        {
            var conservativeMarkdown = MarkdownContentGuard.BuildConservativeMarkdown(options.SourceText);
            var conservativeResult = MarkdownContentGuard.RepairMarkdown(conservativeMarkdown, options.SourceText);
            if (conservativeResult.IsComplete)
            {
                HtmlToMarkdownLog.Debug("Using conservative Markdown derived from the complete source text.");
                return conservativeResult.Markdown;
            }
        }

        return string.IsNullOrWhiteSpace(currentMarkdown)
            ? TryBuildImageOnlyMarkdown(options.FallbackImagePng)
            : currentMarkdown;
    }

    private static string TryBuildImageOnlyMarkdown(byte[]? fallbackImagePng)
    {
        if (fallbackImagePng is not { Length: > 0 })
        {
            return string.Empty;
        }

        var dataUri = $"data:image/png;base64,{System.Convert.ToBase64String(fallbackImagePng)}";
        return $"![clipboard image]({dataUri})";
    }
}
