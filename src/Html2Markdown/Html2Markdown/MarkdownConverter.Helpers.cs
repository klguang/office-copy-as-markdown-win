using System.Globalization;
using System.Net;
using HtmlAgilityPack;

namespace Html2Markdown;

internal static partial class MarkdownConverter
{
    internal static bool TryGetSemanticHeadingLevel(HtmlNode node, IHtmlDialectAdapter dialectAdapter, out int level)
    {
        if (node.Name.Length == 2 &&
            node.Name[0] == 'h' &&
            char.IsDigit(node.Name[1]) &&
            int.TryParse(node.Name[1].ToString(), out level))
        {
            return true;
        }

        return dialectAdapter.TryGetSemanticHeadingLevel(node, out level);
    }

    internal static bool IsBlockContainer(HtmlNode node) =>
        node.Name is "p" or "div" or "section" or "article" or "header" or "footer" or "main" or "li";

    internal static bool IsCandidateBlock(HtmlNode node) =>
        node.Name is "p" or "div" or "section" or "article" or "header" or "footer" or "main";

    internal static bool IsStructuralContainer(HtmlNode node)
    {
        if (node.Name is not ("div" or "section" or "article" or "header" or "footer" or "main"))
        {
            return false;
        }

        return node.ChildNodes.Any(child =>
            child.NodeType == HtmlNodeType.Element &&
            (IsBlockContainer(child) ||
             child.Name is "ul" or "ol" or "table" or "blockquote" or "pre" or "hr" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6"));
    }

    internal static bool IsCodeBlock(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        return style.Contains("white-space:pre", StringComparison.OrdinalIgnoreCase) ||
            (HasMonospaceStyle(node) && node.InnerText.Contains('\n'));
    }

    private static bool IsInlineCode(HtmlNode node) =>
        HasMonospaceStyle(node) && !node.InnerText.Contains('\n');

    private static bool IsBoldElement(HtmlNode node) =>
        node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase) ||
        node.Name.Equals("b", StringComparison.OrdinalIgnoreCase) ||
        HasBoldStyle(node);

    private static bool HasBoldStyle(HtmlNode node) =>
        node.GetAttributeValue("style", string.Empty).Contains("font-weight:bold", StringComparison.OrdinalIgnoreCase);

    private static bool HasItalicStyle(HtmlNode node) =>
        node.GetAttributeValue("style", string.Empty).Contains("font-style:italic", StringComparison.OrdinalIgnoreCase);

    private static bool HasStrikeStyle(HtmlNode node) =>
        node.GetAttributeValue("style", string.Empty).Contains("line-through", StringComparison.OrdinalIgnoreCase);

    private static bool HasMonospaceStyle(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        return style.Contains("Consolas", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("Courier", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("monospace", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsQuoteBlock(HtmlNode node, IHtmlDialectAdapter dialectAdapter) =>
        node.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase) ||
        dialectAdapter.IsQuoteBlock(node);

    internal static bool ShouldSkip(HtmlNode node) =>
        node.NodeType == HtmlNodeType.Comment || node.Name is "script" or "style" or "#document";

    internal static string ApplyQuotePrefix(string text, int quoteDepth)
    {
        if (quoteDepth <= 0)
        {
            return text;
        }

        var prefix = string.Join(' ', Enumerable.Repeat(">", quoteDepth));
        var lines = text.ReplaceLineEndings("\n").Split('\n');
        return string.Join("\n", lines.Select(line => string.IsNullOrWhiteSpace(line) ? prefix : $"{prefix} {line}"));
    }

    internal static string IndentNestedListBlock(string text, int quoteDepth, int indentLevel)
    {
        var indent = new string(' ', indentLevel * 2);
        var lines = text.ReplaceLineEndings("\n").Split('\n');
        return string.Join("\n", lines.Select(line => string.IsNullOrWhiteSpace(line) ? line : $"{indent}{line}"));
    }

    private static bool IsHorizontalRuleText(string text)
    {
        var trimmed = text.Trim();
        return trimmed.Length >= 3 && trimmed.All(ch => ch is '-' or '_' or '*');
    }

    private static string NormalizeTableCell(string text)
    {
        return text.Replace("\r", string.Empty)
            .Replace("\n", "<br>")
            .Replace("|", "\\|")
            .Trim();
    }

    internal static bool TryExtractFontSize(HtmlNode node, out double fontSizePt)
    {
        if (TryParseFontSizeFromStyle(node.GetAttributeValue("style", string.Empty), out fontSizePt))
        {
            return true;
        }

        var descendantCandidates = node.Descendants()
            .Where(descendant => descendant.NodeType == HtmlNodeType.Element)
            .Select(descendant => new
            {
                Node = descendant,
                TextLength = NormalizeInlineText(WebUtility.HtmlDecode(descendant.InnerText)).Length
            })
            .Where(candidate => candidate.TextLength > 0)
            .OrderByDescending(candidate => candidate.TextLength)
            .ToList();

        foreach (var candidate in descendantCandidates)
        {
            if (TryParseFontSizeFromStyle(candidate.Node.GetAttributeValue("style", string.Empty), out fontSizePt))
            {
                return true;
            }
        }

        fontSizePt = 0;
        return false;
    }

    private static bool TryParseFontSizeFromStyle(string style, out double fontSizePt)
    {
        var match = FontSizeRegex.Match(style);
        if (!match.Success)
        {
            fontSizePt = 0;
            return false;
        }

        var value = double.Parse(match.Groups["value"].Value, CultureInfo.InvariantCulture);
        var unit = match.Groups["unit"].Value.ToLowerInvariant();
        fontSizePt = unit switch
        {
            "pt" => value,
            "px" => value * 72d / 96d,
            _ => 0
        };

        return fontSizePt > 0;
    }

    internal static string ExtractCandidateText(HtmlNode node)
    {
        var text = NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText));
        return text.Trim();
    }

    internal static bool IsEntireTextBold(HtmlNode node)
    {
        var textNodes = node
            .DescendantsAndSelf()
            .Where(candidate => candidate.NodeType == HtmlNodeType.Text)
            .Where(candidate => !string.IsNullOrWhiteSpace(NormalizeInlineText(WebUtility.HtmlDecode(candidate.InnerText))))
            .ToList();

        if (textNodes.Count == 0)
        {
            return false;
        }

        foreach (var textNode in textNodes)
        {
            var isBold = false;
            for (var current = textNode.ParentNode; current is not null; current = current.ParentNode)
            {
                if (IsBoldElement(current))
                {
                    isBold = true;
                    break;
                }

                if (current == node)
                {
                    break;
                }
            }

            if (!isBold)
            {
                return false;
            }
        }

        return true;
    }

    internal static bool HasBlockedAncestor(HtmlNode node, IHtmlDialectAdapter dialectAdapter)
    {
        for (var current = node.ParentNode; current is not null; current = current.ParentNode)
        {
            if (current.Name is "li" or "td" or "th" or "blockquote" or "pre")
            {
                return true;
            }

            if (current.Name is "ul" or "ol" or "table")
            {
                return true;
            }

            if (IsQuoteBlock(current, dialectAdapter) || IsCodeBlock(current) || dialectAdapter.ShouldBlockHeadingInference(current))
            {
                return true;
            }
        }

        return false;
    }
}
