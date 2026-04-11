using System.Text.RegularExpressions;

namespace OfficeCopyAsMarkdown.Services;

internal static partial class MarkdownConverter
{
    private const int MaximumCandidateHeadingLength = 30;
    private static readonly CandidateHeadingInferenceOptions CandidateHeadingInference = new(MaxLevels: 4, SparseStartLevel: 2);
    private static readonly char[] DisallowedHeadingTerminators = ['\u3002', '.', '\uFF0C', ',', '\uFF1B', ';'];
    private static readonly Regex HeadingStyleRegex = new(@"mso-style-name:\s*[""']?Heading\s*(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex HeadingClassRegex = new(@"Heading(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex OfficeLevelRegex = new(@"level(?<level>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex FontSizeRegex = new(@"font-size\s*:\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>pt|px)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex NumberedHeadingRegex = new(@"^(?<prefix>\d+(?:\.\d+)+)(?<rest>\S)", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldSegmentsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<middle>[:/\-\uFF1A])\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldRunsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex BoldLabelSeparatedColonRegex = new(@"\*\*(?<label>[^*\r\n]+)\*\*(?<colon>[:\uFF1A])(?<next>\S)", RegexOptions.Compiled);
    private static readonly Regex BoldLabelFollowedByTextRegex = new(@"\*\*(?<label>[^*\r\n]+)\*\*(?<next>\S)", RegexOptions.Compiled);
    private static readonly Regex TextFollowedByBoldLabelRegex = new(@"(?<prev>[\p{L}\p{N}\u4E00-\u9FFF])\*\*(?<label>[^*\r\n]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex TightPlusBetweenWordsRegex = new(@"(?<left>[\p{L}\p{N}\u4E00-\u9FFF])\+(?<right>[\p{L}\p{N}\u4E00-\u9FFF])", RegexOptions.Compiled);

    public static string Convert(string html, byte[]? fallbackImagePng)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var normalized = OfficeHtmlNormalizer.Normalize(html);
        var document = new HtmlAgilityPack.HtmlDocument();
        document.LoadHtml(normalized);

        var root = document.DocumentNode.SelectSingleNode("//body") ?? document.DocumentNode;
        var context = new ConversionContext(fallbackImagePng, root.SelectNodes(".//img")?.Count ?? 0);
        var headingInference = HeadingInference.Analyze(root, CandidateHeadingInference);
        var blocks = ConvertBlocks(root.ChildNodes, context, headingInference, 0);
        var markdown = string.Join("\n\n", blocks.Where(block => !string.IsNullOrWhiteSpace(block))).Trim();

        if (string.IsNullOrWhiteSpace(markdown) && fallbackImagePng is { Length: > 0 })
        {
            var dataUri = $"data:image/png;base64,{System.Convert.ToBase64String(fallbackImagePng)}";
            return $"![clipboard image]({dataUri})";
        }

        return markdown;
    }

    private sealed record CandidateHeadingInferenceOptions(int MaxLevels, int SparseStartLevel)
    {
        public int EffectiveMaxLevels => Math.Max(1, MaxLevels);
        public int EffectiveSparseStartLevel => Math.Clamp(SparseStartLevel, 1, 6);
    }
}
