using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using Html2Markdown;
using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Services;

internal sealed class OfficeHtmlDialectAdapter : IHtmlDialectAdapter
{
    private static readonly Regex HeadingStyleRegex = new(@"mso-style-name:\s*[""']?Heading\s*(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex HeadingClassRegex = new(@"Heading(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex OfficeLevelRegex = new(@"level(?<level>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static OfficeHtmlDialectAdapter Instance { get; } = new();

    private OfficeHtmlDialectAdapter()
    {
    }

    public bool TryGetSemanticHeadingLevel(HtmlNode node, out int level)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        var className = node.GetAttributeValue("class", string.Empty);
        var match = HeadingStyleRegex.Match(style);
        if (!match.Success)
        {
            match = HeadingClassRegex.Match(className);
        }

        if (match.Success && int.TryParse(match.Groups["level"].Value, out level))
        {
            return true;
        }

        level = 0;
        return false;
    }

    public bool IsQuoteBlock(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        return style.Contains("mso-style-name:Quote", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("border-left", StringComparison.OrdinalIgnoreCase);
    }

    public bool TryConvertListLikeBlock(HtmlNode node, int quoteDepth, out string markdown)
    {
        markdown = string.Empty;

        if (!IsOfficeListParagraph(node))
        {
            return false;
        }

        var rawText = NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText));
        if (string.IsNullOrWhiteSpace(rawText))
        {
            return false;
        }

        var style = node.GetAttributeValue("style", string.Empty);
        var levelMatch = OfficeLevelRegex.Match(style);
        var indentLevel = levelMatch.Success ? Math.Max(0, int.Parse(levelMatch.Groups["level"].Value, CultureInfo.InvariantCulture) - 1) : 0;
        var indent = new string(' ', indentLevel * 2);

        if (TryConvertTaskLine(rawText, out var taskLine))
        {
            markdown = ApplyQuotePrefix($"{indent}{taskLine}", quoteDepth);
            return true;
        }

        if (TryConvertLooseOrderedLine(rawText, out var orderedLine))
        {
            markdown = ApplyQuotePrefix($"{indent}{orderedLine}", quoteDepth);
            return true;
        }

        if (TryConvertLooseBulletLine(rawText, out var bulletLine))
        {
            markdown = ApplyQuotePrefix($"{indent}{bulletLine}", quoteDepth);
            return true;
        }

        markdown = ApplyQuotePrefix($"{indent}- {rawText}", quoteDepth);
        return true;
    }

    public bool ShouldBlockHeadingInference(HtmlNode node) =>
        IsOfficeListParagraph(node) || IsQuoteBlock(node);

    private static bool IsOfficeListParagraph(HtmlNode node)
    {
        if (!node.Name.Equals("p", StringComparison.OrdinalIgnoreCase) &&
            !node.Name.Equals("div", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var style = node.GetAttributeValue("style", string.Empty);
        var className = node.GetAttributeValue("class", string.Empty);
        return style.Contains("mso-list", StringComparison.OrdinalIgnoreCase) ||
            className.Contains("MsoListParagraph", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeInlineText(string text)
    {
        var normalized = text
            .Replace('\u00A0', ' ')
            .Replace('\u200B', ' ');

        return MarkdownLineSyntax.NormalizeWhitespace(normalized);
    }

    private static string ApplyQuotePrefix(string text, int quoteDepth)
    {
        if (quoteDepth <= 0)
        {
            return text;
        }

        var prefix = string.Join(' ', Enumerable.Repeat(">", quoteDepth));
        var lines = text.ReplaceLineEndings("\n").Split('\n');
        return string.Join("\n", lines.Select(line => string.IsNullOrWhiteSpace(line) ? prefix : $"{prefix} {line}"));
    }

    private static bool TryConvertTaskLine(string text, out string markdown)
    {
        if (MarkdownLineSyntax.TryParseSourceTask(text, out var content, out var isChecked))
        {
            markdown = MarkdownLineSyntax.RenderMarkdownTask(content, isChecked);
            return true;
        }

        markdown = string.Empty;
        return false;
    }

    private static bool TryConvertLooseOrderedLine(string text, out string markdown)
    {
        if (MarkdownLineSyntax.TryParseSourceOrderedListItem(text, out var marker, out var content))
        {
            markdown = MarkdownLineSyntax.RenderMarkdownOrderedListItem(marker, content);
            return true;
        }

        markdown = string.Empty;
        return false;
    }

    private static bool TryConvertLooseBulletLine(string text, out string markdown)
    {
        if (MarkdownLineSyntax.TryParseSourceBulletListItem(text, out var content))
        {
            markdown = MarkdownLineSyntax.RenderMarkdownBulletListItem(content);
            return true;
        }

        markdown = string.Empty;
        return false;
    }
}
