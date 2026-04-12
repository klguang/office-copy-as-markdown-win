using System.Text.RegularExpressions;

namespace Html2Markdown;

internal static partial class MarkdownConverter
{
    private static readonly Regex FontSizeRegex = new(@"font-size\s*:\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>pt|px)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex NumberedHeadingRegex = new(@"^(?<prefix>\d+(?:\.\d+)+)(?<rest>\S)", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldSegmentsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<middle>[:/\-\uFF1A])\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldRunsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex BoldLabelSeparatedColonRegex = new(@"\*\*(?<label>[^*\r\n]+)\*\*(?<colon>[:\uFF1A])(?<next>\S)", RegexOptions.Compiled);
    private static readonly Regex BoldLabelFollowedByTextRegex = new(@"\*\*(?<label>[^*\r\n]+)\*\*(?<next>\S)", RegexOptions.Compiled);
    private static readonly Regex TextFollowedByBoldLabelRegex = new(@"(?<prev>[\p{L}\p{N}\u4E00-\u9FFF])\*\*(?<label>[^*\r\n]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex TightPlusBetweenWordsRegex = new(@"(?<left>[\p{L}\p{N}\u4E00-\u9FFF])\+(?<right>[\p{L}\p{N}\u4E00-\u9FFF])", RegexOptions.Compiled);

    public static string Convert(
        string html,
        byte[]? fallbackImagePng,
        IHtmlDialectAdapter dialectAdapter,
        IHeadingInferenceStrategy headingInferenceStrategy)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var document = new HtmlAgilityPack.HtmlDocument();
        document.LoadHtml(html);

        var root = document.DocumentNode.SelectSingleNode("//body") ?? document.DocumentNode;
        var context = new ConversionContext(fallbackImagePng, dialectAdapter);
        var inferredHeadingLevels = headingInferenceStrategy.InferHeadingLevels(root, dialectAdapter);
        var blocks = ConvertBlocks(root.ChildNodes, context, inferredHeadingLevels, 0);
        var markdown = string.Join("\n\n", blocks.Where(block => !string.IsNullOrWhiteSpace(block))).Trim();

        if (string.IsNullOrWhiteSpace(markdown) && fallbackImagePng is { Length: > 0 })
        {
            var dataUri = $"data:image/png;base64,{System.Convert.ToBase64String(fallbackImagePng)}";
            return $"![clipboard image]({dataUri})";
        }

        return markdown;
    }
}
