using System.Text.RegularExpressions;

namespace Html2Markdown;

public static class MarkdownLineSyntax
{
    private static readonly Regex MarkdownPrefixRegex = new(@"^(>+\s*)?(#{1,6}\s+)?((\d+|[A-Za-z]+)[\.\)]\s+|[-*+]\s+)?(\[[ xX]\]\s+)?", RegexOptions.Compiled);
    private static readonly Regex SourcePrefixRegex = new(@"^((\d+|[A-Za-z]+)[\.\)]\s+|[\u2022\u00B7\u25CB\u25E6\u25CFo\-*]\s+|[\u2610\u2611\u2612]\s+)+", RegexOptions.Compiled);
    private static readonly Regex OrderedListRegex = new(@"^(?<marker>(\d+|[A-Za-z]+)[\.\)])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletListRegex = new(@"^(?<marker>[\u2022\u00B7\u25CB\u25E6\u25CFo\-*])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex CheckboxRegex = new(@"^(?<box>[\u2610\u2611\u2612])\s*(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    public static string NormalizeWhitespace(string text) =>
        WhitespaceRegex.Replace(text, " ");

    public static string StripMarkdownPrefix(string text) =>
        MarkdownPrefixRegex.Replace(text, string.Empty);

    public static string StripSourcePrefix(string text) =>
        SourcePrefixRegex.Replace(text, string.Empty);

    public static bool TryParseSourceTask(string text, out string content, out bool isChecked)
    {
        var match = CheckboxRegex.Match(text);
        if (match.Success)
        {
            content = match.Groups["text"].Value.Trim();
            isChecked = match.Groups["box"].Value is "\u2611" or "\u2612";
            return true;
        }

        content = string.Empty;
        isChecked = false;
        return false;
    }

    public static bool TryParseSourceOrderedListItem(string text, out string marker, out string content)
    {
        var match = OrderedListRegex.Match(text);
        if (match.Success)
        {
            marker = match.Groups["marker"].Value;
            content = match.Groups["text"].Value.Trim();
            return true;
        }

        marker = string.Empty;
        content = string.Empty;
        return false;
    }

    public static bool TryParseSourceBulletListItem(string text, out string content)
    {
        var match = BulletListRegex.Match(text);
        if (match.Success)
        {
            content = match.Groups["text"].Value.Trim();
            return true;
        }

        content = string.Empty;
        return false;
    }

    public static bool LooksLikeOrderedListItem(string text) =>
        OrderedListRegex.IsMatch(text.TrimStart());

    public static bool LooksLikeBulletListItem(string text) =>
        BulletListRegex.IsMatch(text.TrimStart());

    public static bool LooksLikeTaskListItem(string text) =>
        text.TrimStart().StartsWith("- [", StringComparison.Ordinal);

    public static string RenderMarkdownTaskMarker(bool isChecked) =>
        isChecked ? "[x]" : "[ ]";

    public static string RenderMarkdownTask(string content, bool isChecked) =>
        $"- {RenderMarkdownTaskMarker(isChecked)} {content}";

    public static string RenderMarkdownOrderedListItem(string marker, string content) =>
        $"{marker} {content}";

    public static string RenderMarkdownBulletListItem(string content) =>
        $"- {content}";
}
