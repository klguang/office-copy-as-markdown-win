using System.Net;
using System.Text;
using HtmlAgilityPack;

namespace Html2Markdown;

internal static class HtmlStructureLogger
{
    private const int MaximumNodesToLog = 400;
    private const int MaximumTextLength = 80;
    private const int MaximumAttributeLength = 120;

    public static void LogFragmentStructure(string html)
    {
        if (string.IsNullOrWhiteSpace(html) || !AppLogger.IsEnabled || AppLogger.CurrentLevel < LogLevel.Debug)
        {
            return;
        }

        try
        {
            var document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(html);

            var root = document.DocumentNode.SelectSingleNode("//body") ?? document.DocumentNode;
            var builder = new StringBuilder();
            builder.AppendLine("HTML fragment structure:");

            var loggedNodes = 0;
            foreach (var child in root.ChildNodes)
            {
                AppendNode(builder, child, depth: 0, ref loggedNodes);
                if (loggedNodes >= MaximumNodesToLog)
                {
                    builder.AppendLine("... structure log truncated ...");
                    break;
                }
            }

            AppLogger.Debug(builder.ToString().TrimEnd());
        }
        catch (Exception ex)
        {
            AppLogger.Warning($"Failed to log HTML fragment structure: {ex.Message}");
        }
    }

    private static void AppendNode(StringBuilder builder, HtmlNode node, int depth, ref int loggedNodes)
    {
        if (loggedNodes >= MaximumNodesToLog || ShouldSkip(node))
        {
            return;
        }

        loggedNodes++;
        var indent = new string(' ', depth * 2);

        if (node.NodeType == HtmlNodeType.Text)
        {
            var text = NormalizePreview(WebUtility.HtmlDecode(node.InnerText));
            if (!string.IsNullOrWhiteSpace(text))
            {
                builder.Append(indent)
                    .Append("- text: ")
                    .AppendLine(TrimToLength(text, MaximumTextLength));
            }

            return;
        }

        builder.Append(indent)
            .Append("- <")
            .Append(node.Name);

        var className = node.GetAttributeValue("class", string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(className))
        {
            builder.Append(" class=\"")
                .Append(TrimToLength(className, MaximumAttributeLength))
                .Append('"');
        }

        var style = node.GetAttributeValue("style", string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(style))
        {
            builder.Append(" style=\"")
                .Append(TrimToLength(NormalizePreview(style), MaximumAttributeLength))
                .Append('"');
        }

        var textPreview = NormalizePreview(WebUtility.HtmlDecode(node.InnerText));
        if (!string.IsNullOrWhiteSpace(textPreview))
        {
            builder.Append(" text=\"")
                .Append(TrimToLength(textPreview, MaximumTextLength))
                .Append('"');
        }

        builder.AppendLine(">");

        foreach (var child in node.ChildNodes)
        {
            AppendNode(builder, child, depth + 1, ref loggedNodes);
            if (loggedNodes >= MaximumNodesToLog)
            {
                return;
            }
        }
    }

    private static bool ShouldSkip(HtmlNode node) =>
        node.NodeType == HtmlNodeType.Comment || node.Name is "script" or "style";

    private static string NormalizePreview(string value) =>
        value.ReplaceLineEndings(" ").Replace('\u00A0', ' ').Trim();

    private static string TrimToLength(string value, int maxLength) =>
        value.Length <= maxLength ? value : $"{value[..maxLength]}...";
}
