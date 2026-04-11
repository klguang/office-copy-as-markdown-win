using System.Globalization;
using System.Net;
using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Services;

internal static partial class MarkdownConverter
{
    private static bool TryConvertOfficeListParagraph(HtmlNode node, int quoteDepth, out string markdown)
    {
        markdown = string.Empty;

        if (!node.Name.Equals("p", StringComparison.OrdinalIgnoreCase) &&
            !node.Name.Equals("div", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var style = node.GetAttributeValue("style", string.Empty);
        var className = node.GetAttributeValue("class", string.Empty);
        if (!style.Contains("mso-list", StringComparison.OrdinalIgnoreCase) &&
            !className.Contains("MsoListParagraph", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var rawText = NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText));
        if (string.IsNullOrWhiteSpace(rawText))
        {
            return false;
        }

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

    private static bool TryGetTaskMarker(HtmlNode liNode, out string marker)
    {
        var checkbox = liNode.SelectSingleNode(".//input[@type='checkbox']");
        if (checkbox is not null)
        {
            marker = ConvertCheckbox(checkbox);
            return true;
        }

        var text = NormalizeInlineText(WebUtility.HtmlDecode(liNode.InnerText));
        if (TryConvertTaskLine(text, out var taskLine))
        {
            marker = taskLine.Split(' ', 3)[1];
            return true;
        }

        marker = string.Empty;
        return false;
    }

    private static string ConvertCheckbox(HtmlNode checkboxNode)
    {
        var isChecked = checkboxNode.Attributes["checked"] is not null ||
            checkboxNode.GetAttributeValue("aria-checked", string.Empty).Equals("true", StringComparison.OrdinalIgnoreCase);

        return MarkdownLineSyntax.RenderMarkdownTaskMarker(isChecked);
    }
}
