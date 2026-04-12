using System.Globalization;
using System.Net;
using HtmlAgilityPack;

namespace Html2Markdown;

internal static partial class MarkdownConverter
{
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
