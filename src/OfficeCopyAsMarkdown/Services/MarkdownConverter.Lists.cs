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
        var match = CheckboxRegex.Match(text);
        if (match.Success)
        {
            var marker = match.Groups["box"].Value is "\u2611" or "\u2612" ? "[x]" : "[ ]";
            markdown = $"- {marker} {match.Groups["text"].Value.Trim()}";
            return true;
        }

        markdown = string.Empty;
        return false;
    }

    private static bool TryConvertLooseOrderedLine(string text, out string markdown)
    {
        var match = OrderedListRegex.Match(text);
        if (match.Success)
        {
            markdown = $"{match.Groups["marker"].Value} {match.Groups["text"].Value.Trim()}";
            return true;
        }

        markdown = string.Empty;
        return false;
    }

    private static bool TryConvertLooseBulletLine(string text, out string markdown)
    {
        var match = BulletListRegex.Match(text);
        if (match.Success)
        {
            markdown = $"- {match.Groups["text"].Value.Trim()}";
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

        return isChecked ? "[x]" : "[ ]";
    }
}
