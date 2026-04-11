using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Services;

internal static class MarkdownConverter
{
    private const int MaximumCandidateHeadingLength = 20;
    private static readonly char[] DisallowedHeadingTerminators = ['\u3002', '.', '\uFF0C', ',', '\uFF1B', ';'];
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex CheckboxRegex = new(@"^(?<box>[\u2610\u2611\u2612])\s*(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex OrderedListRegex = new(@"^(?<marker>(\d+|[A-Za-z]+)[\.\)])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletListRegex = new(@"^(?<marker>[\u2022\u00B7\u25CB\u25E6\u25CFo\-*])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex HeadingStyleRegex = new(@"mso-style-name:\s*[""']?Heading\s*(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex HeadingClassRegex = new(@"Heading(?<level>\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex OfficeLevelRegex = new(@"level(?<level>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex FontSizeRegex = new(@"font-size\s*:\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>pt|px)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex NumberedHeadingRegex = new(@"^(?<prefix>\d+(?:\.\d+)+)(?<rest>\S)", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldSegmentsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<middle>[:/\-\uFF1A])\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);
    private static readonly Regex AdjacentBoldRunsRegex = new(@"\*\*(?<left>[^*]+)\*\*\s*\*\*(?<right>[^*]+)\*\*", RegexOptions.Compiled);

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
        var headingInference = HeadingInference.Analyze(root);
        var blocks = ConvertBlocks(root.ChildNodes, context, headingInference, 0);
        var markdown = string.Join("\n\n", blocks.Where(block => !string.IsNullOrWhiteSpace(block))).Trim();

        if (string.IsNullOrWhiteSpace(markdown) && fallbackImagePng is { Length: > 0 })
        {
            var dataUri = $"data:image/png;base64,{System.Convert.ToBase64String(fallbackImagePng)}";
            return $"![clipboard image]({dataUri})";
        }

        return markdown;
    }

    private static List<string> ConvertBlocks(HtmlNodeCollection nodes, ConversionContext context, HeadingInference headingInference, int quoteDepth)
    {
        var blocks = new List<string>();

        for (var index = 0; index < nodes.Count; index++)
        {
            var node = nodes[index];
            if (ShouldSkip(node))
            {
                continue;
            }

            if (TryConvertOfficeListParagraph(node, quoteDepth, out var officeList))
            {
                blocks.Add(officeList);
                continue;
            }

            if (IsStructuralContainer(node))
            {
                var nestedBlocks = ConvertBlocks(node.ChildNodes, context, headingInference, quoteDepth);
                if (nestedBlocks.Count > 0)
                {
                    blocks.AddRange(nestedBlocks);
                }

                continue;
            }

            if (node.Name is "ul" or "ol")
            {
                blocks.Add(ConvertList(node, context, headingInference, quoteDepth, 0));
                continue;
            }

            if (node.Name.Equals("table", StringComparison.OrdinalIgnoreCase))
            {
                blocks.Add(ApplyQuotePrefix(ConvertTable(node, context), quoteDepth));
                continue;
            }

            if (node.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase) || HasQuoteStyle(node))
            {
                var nested = ConvertBlocks(node.ChildNodes, context, headingInference, quoteDepth + 1);
                if (nested.Count > 0)
                {
                    blocks.Add(string.Join("\n\n", nested));
                }

                continue;
            }

            if (node.Name.Equals("pre", StringComparison.OrdinalIgnoreCase) || IsCodeBlock(node))
            {
                blocks.Add(ApplyQuotePrefix(ConvertCodeBlock(node), quoteDepth));
                continue;
            }

            if (node.Name.Equals("hr", StringComparison.OrdinalIgnoreCase))
            {
                blocks.Add(ApplyQuotePrefix("---", quoteDepth));
                continue;
            }

            if (TryGetSemanticHeadingLevel(node, out var headingLevel) ||
                headingInference.TryGetLevel(node, out headingLevel))
            {
                var headingText = NormalizeHeadingText(ConvertInlines(node, context).Trim());
                if (!string.IsNullOrWhiteSpace(headingText))
                {
                    blocks.Add(ApplyQuotePrefix($"{new string('#', headingLevel)} {headingText}", quoteDepth));
                }

                continue;
            }

            if (node.Name.Equals("img", StringComparison.OrdinalIgnoreCase))
            {
                var image = ConvertImage(node, context);
                if (!string.IsNullOrWhiteSpace(image))
                {
                    blocks.Add(ApplyQuotePrefix(image, quoteDepth));
                }

                continue;
            }

            if (IsBlockContainer(node))
            {
                blocks.AddRange(ConvertTextBlocks(ConvertInlines(node, context), quoteDepth, node, headingInference));
                continue;
            }

            if (TryConvertInlineSequence(nodes, ref index, context, headingInference, quoteDepth, out var inlineBlocks))
            {
                blocks.AddRange(inlineBlocks);
            }
        }

        return blocks;
    }

    private static string ConvertList(HtmlNode listNode, ConversionContext context, HeadingInference headingInference, int quoteDepth, int indentLevel)
    {
        var isOrdered = listNode.Name.Equals("ol", StringComparison.OrdinalIgnoreCase);
        var itemIndex = 1;
        var lines = new List<string>();
        var hasRenderedListItem = false;

        foreach (var child in listNode.ChildNodes)
        {
            if (ShouldSkip(child))
            {
                continue;
            }

            if (child.NodeType == HtmlNodeType.Text && string.IsNullOrWhiteSpace(WebUtility.HtmlDecode(child.InnerText)))
            {
                continue;
            }

            if (child.Name.Equals("li", StringComparison.OrdinalIgnoreCase))
            {
                var nestedLists = child.ChildNodes.Where(node => node.Name is "ul" or "ol").ToList();
                var inlineNodes = child.ChildNodes.Where(node => node.Name is not "ul" and not "ol").ToList();
                var indent = new string(' ', indentLevel * 2);
                var prefix = isOrdered ? $"{itemIndex}. " : "- ";

                if (TryGetTaskMarker(child, out var taskMarker))
                {
                    prefix = $"- {taskMarker} ";
                }

                var content = ConvertNodesAsInline(inlineNodes, context).Trim();
                lines.Add(ApplyQuotePrefix($"{indent}{prefix}{content}".TrimEnd(), quoteDepth));

                var nestedBlockContainers = child.ChildNodes
                    .Where(node => node.Name is not "ul" and not "ol" && node.NodeType == HtmlNodeType.Element)
                    .ToList();

                foreach (var nested in nestedLists)
                {
                    lines.Add(ConvertList(nested, context, headingInference, quoteDepth, indentLevel + 1));
                }

                foreach (var nestedBlock in nestedBlockContainers)
                {
                    if (!IsStructuralContainer(nestedBlock))
                    {
                        continue;
                    }

                    var nestedBlocks = ConvertBlocks(nestedBlock.ChildNodes, context, headingInference, 0);
                    foreach (var nestedLine in nestedBlocks.Where(line => !string.IsNullOrWhiteSpace(line)))
                    {
                        lines.Add(IndentNestedListBlock(nestedLine, quoteDepth, indentLevel + 1));
                    }
                }

                itemIndex++;
                hasRenderedListItem = true;
                continue;
            }

            if (child.Name is "ul" or "ol")
            {
                var nestedIndentLevel = hasRenderedListItem ? indentLevel + 1 : indentLevel;
                var nestedList = ConvertList(child, context, headingInference, quoteDepth, nestedIndentLevel);
                if (!string.IsNullOrWhiteSpace(nestedList))
                {
                    lines.Add(nestedList);
                }

                continue;
            }

            if (IsStructuralContainer(child))
            {
                var nestedIndentLevel = hasRenderedListItem ? indentLevel + 1 : indentLevel;
                var nestedBlocks = ConvertBlocks(child.ChildNodes, context, headingInference, 0);
                foreach (var nestedLine in nestedBlocks.Where(line => !string.IsNullOrWhiteSpace(line)))
                {
                    lines.Add(IndentNestedListBlock(nestedLine, quoteDepth, nestedIndentLevel));
                }

                continue;
            }

            var fallbackText = ConvertNodesAsInline([child], context).Trim();
            if (!string.IsNullOrWhiteSpace(fallbackText))
            {
                var nestedIndentLevel = hasRenderedListItem ? indentLevel + 1 : indentLevel;
                lines.Add(IndentNestedListBlock(fallbackText, quoteDepth, nestedIndentLevel));
            }
        }

        return string.Join("\n", lines.Where(line => !string.IsNullOrWhiteSpace(line)));
    }

    private static string ConvertTable(HtmlNode tableNode, ConversionContext context)
    {
        var rows = tableNode.SelectNodes(".//tr");
        if (rows is null || rows.Count == 0)
        {
            return ConvertInlines(tableNode, context);
        }

        var parsedRows = rows
            .Select(row => row.Elements("th").Concat(row.Elements("td")).ToList())
            .Where(row => row.Count > 0)
            .ToList();

        if (parsedRows.Count == 0)
        {
            return string.Empty;
        }

        var header = parsedRows[0].Select(cell => NormalizeTableCell(ConvertInlines(cell, context))).ToList();
        var builder = new StringBuilder();
        builder.Append("| ");
        builder.Append(string.Join(" | ", header));
        builder.AppendLine(" |");
        builder.Append("| ");
        builder.Append(string.Join(" | ", header.Select(_ => "---")));
        builder.AppendLine(" |");

        foreach (var row in parsedRows.Skip(1))
        {
            builder.Append("| ");
            builder.Append(string.Join(" | ", row.Select(cell => NormalizeTableCell(ConvertInlines(cell, context)))));
            builder.AppendLine(" |");
        }

        return builder.ToString().TrimEnd();
    }

    private static string ConvertCodeBlock(HtmlNode node)
    {
        var text = WebUtility.HtmlDecode(node.InnerText)
            .ReplaceLineEndings("\n")
            .Trim('\n');

        return $"```\n{text}\n```";
    }

    private static string ConvertInlines(HtmlNode node, ConversionContext context)
    {
        if (node.Name.Equals("pre", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertCodeBlock(node);
        }

        return ConvertNodesAsInline(node.ChildNodes, context);
    }

    private static string ConvertNodesAsInline(IEnumerable<HtmlNode> nodes, ConversionContext context)
    {
        var builder = new StringBuilder();

        foreach (var node in nodes)
        {
            if (ShouldSkip(node))
            {
                continue;
            }

            if (node.NodeType == HtmlNodeType.Text)
            {
                builder.Append(NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText)));
                continue;
            }

            builder.Append(ConvertInlineElement(node, context));
        }

        return NormalizeInlineResult(builder.ToString());
    }

    private static string ConvertInlineElement(HtmlNode node, ConversionContext context)
    {
        if (node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
        {
            return "  \n";
        }

        if (node.Name.Equals("a", StringComparison.OrdinalIgnoreCase))
        {
            var text = ConvertNodesAsInline(node.ChildNodes, context).Trim();
            var href = node.GetAttributeValue("href", string.Empty).Trim();
            return string.IsNullOrWhiteSpace(href) ? text : $"[{text}]({href})";
        }

        if (node.Name.Equals("img", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertImage(node, context);
        }

        if (node.Name.Equals("code", StringComparison.OrdinalIgnoreCase) || IsInlineCode(node))
        {
            var inlineCode = ConvertNodesAsInline(node.ChildNodes, context).Trim();
            return string.IsNullOrWhiteSpace(inlineCode) ? string.Empty : WrapInlineCode(inlineCode);
        }

        if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase) || node.Name.Equals("b", StringComparison.OrdinalIgnoreCase))
        {
            return Wrap("**", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name.Equals("em", StringComparison.OrdinalIgnoreCase) || node.Name.Equals("i", StringComparison.OrdinalIgnoreCase))
        {
            return Wrap("*", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name is "del" or "s" or "strike" || HasStrikeStyle(node))
        {
            return Wrap("~~", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name.Equals("input", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertCheckbox(node);
        }

        var content = ConvertNodesAsInline(node.ChildNodes, context);

        if (HasBoldStyle(node))
        {
            content = Wrap("**", content);
        }

        if (HasItalicStyle(node))
        {
            content = Wrap("*", content);
        }

        if (HasStrikeStyle(node))
        {
            content = Wrap("~~", content);
        }

        if (IsInlineCode(node))
        {
            content = WrapInlineCode(content);
        }

        return content;
    }

    private static string ConvertImage(HtmlNode node, ConversionContext context)
    {
        var src = node.GetAttributeValue("src", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(src) && context.CanUseFallbackImage)
        {
            src = context.UseFallbackImage();
        }

        if (string.IsNullOrWhiteSpace(src))
        {
            return string.Empty;
        }

        var alt = node.GetAttributeValue("alt", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(alt))
        {
            alt = "image";
        }

        return $"![{alt}]({src})";
    }

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

    private static bool TryGetSemanticHeadingLevel(HtmlNode node, out int level)
    {
        if (node.Name.Length == 2 &&
            node.Name[0] == 'h' &&
            char.IsDigit(node.Name[1]) &&
            int.TryParse(node.Name[1].ToString(), out level))
        {
            return true;
        }

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

    private static bool IsBlockContainer(HtmlNode node) =>
        node.Name is "p" or "div" or "section" or "article" or "header" or "footer" or "main" or "li";

    private static bool IsCandidateBlock(HtmlNode node) =>
        node.Name is "p" or "div" or "section" or "article" or "header" or "footer" or "main";

    private static bool IsStructuralContainer(HtmlNode node)
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

    private static bool IsCodeBlock(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        return style.Contains("white-space:pre", StringComparison.OrdinalIgnoreCase) ||
            (HasMonospaceStyle(node) && node.InnerText.Contains('\n'));
    }

    private static bool IsInlineCode(HtmlNode node) =>
        HasMonospaceStyle(node) && !node.InnerText.Contains('\n');

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

    private static bool HasQuoteStyle(HtmlNode node)
    {
        var style = node.GetAttributeValue("style", string.Empty);
        return style.Contains("mso-style-name:Quote", StringComparison.OrdinalIgnoreCase) ||
            style.Contains("border-left", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldSkip(HtmlNode node) =>
        node.NodeType == HtmlNodeType.Comment || node.Name is "script" or "style" or "#document";

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

    private static string Wrap(string marker, string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        var leadingWhitespace = content.Length - content.TrimStart().Length;
        var trailingWhitespace = content.Length - content.TrimEnd().Length;
        var trimmed = content.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        return $"{new string(' ', leadingWhitespace)}{marker}{trimmed}{marker}{new string(' ', trailingWhitespace)}";
    }

    private static string WrapInlineCode(string content)
    {
        var trimmed = content.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        return $"`{trimmed.Replace("`", "\\`", StringComparison.Ordinal)}`";
    }

    private static string NormalizeInlineText(string text)
    {
        var normalized = text
            .Replace('\u00A0', ' ')
            .Replace('\u200B', ' ');

        return WhitespaceRegex.Replace(normalized, " ");
    }

    private static string NormalizeInlineResult(string text)
    {
        var normalized = Regex.Replace(text, @"[ \t]+(\r?\n)", "$1").Trim();
        while (true)
        {
            var previous = normalized;
            normalized = AdjacentBoldSegmentsRegex.Replace(normalized, match =>
                $"**{match.Groups["left"].Value}{match.Groups["middle"].Value}{match.Groups["right"].Value}**");
            normalized = AdjacentBoldRunsRegex.Replace(normalized, match =>
            {
                var left = match.Groups["left"].Value;
                var right = match.Groups["right"].Value;
                return $"**{left}{ResolveAdjacentBoldJoiner(left, right)}{right}**";
            });

            if (string.Equals(previous, normalized, StringComparison.Ordinal))
            {
                break;
            }
        }

        return normalized;
    }

    private static string NormalizeHeadingText(string text)
    {
        var normalized = text.Trim();

        normalized = normalized
            .Replace("**", string.Empty, StringComparison.Ordinal)
            .Replace("__", string.Empty, StringComparison.Ordinal);

        normalized = NormalizeInlineText(normalized).Trim();
        normalized = NumberedHeadingRegex.Replace(normalized, "${prefix} ${rest}");
        return normalized;
    }

    private static string IndentNestedListBlock(string text, int quoteDepth, int indentLevel)
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

    private static bool IsDisallowedHeadingTerminator(char character) =>
        DisallowedHeadingTerminators.Contains(character);

    private static string ResolveAdjacentBoldJoiner(string left, string right)
    {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
        {
            return string.Empty;
        }

        var leftChar = left[^1];
        var rightChar = right[0];

        if (char.IsWhiteSpace(leftChar) || char.IsWhiteSpace(rightChar))
        {
            return string.Empty;
        }

        if (IsOpeningPunctuation(leftChar) || IsClosingPunctuation(rightChar))
        {
            return string.Empty;
        }

        if (IsSeparatorPunctuation(leftChar) || IsSeparatorPunctuation(rightChar))
        {
            return string.Empty;
        }

        if ((IsCjk(leftChar) && IsLatinOrDigit(rightChar)) ||
            (IsLatinOrDigit(leftChar) && IsCjk(rightChar)) ||
            (char.IsLetterOrDigit(leftChar) && char.IsLetterOrDigit(rightChar)))
        {
            return " ";
        }

        return string.Empty;
    }

    private static bool IsCjk(char value) => value is >= '\u4E00' and <= '\u9FFF';

    private static bool IsLatinOrDigit(char value) =>
        (value is >= 'A' and <= 'Z') ||
        (value is >= 'a' and <= 'z') ||
        char.IsDigit(value);

    private static bool IsOpeningPunctuation(char value) => value is '(' or '[' or '{' or '<' or '（' or '《' or '“' or '"' or '\'';

    private static bool IsClosingPunctuation(char value) => value is ')' or ']' or '}' or '>' or '）' or '》' or '”' or '"' or '\'';

    private static bool IsSeparatorPunctuation(char value) => value is ':' or '：' or '/' or '\\' or '-' or '—' or '+' or '＋';

    private static bool TryExtractFontSize(HtmlNode node, out double fontSizePt)
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

    private static string ExtractCandidateText(HtmlNode node)
    {
        var text = NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText));
        return text.Trim();
    }

    private static bool HasBlockedAncestor(HtmlNode node)
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

            if (HasQuoteStyle(current) || IsCodeBlock(current))
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryConvertInlineSequence(
        HtmlNodeCollection nodes,
        ref int index,
        ConversionContext context,
        HeadingInference headingInference,
        int quoteDepth,
        out List<string> blocks)
    {
        blocks = [];

        if (!IsInlineSequenceCandidate(nodes[index]))
        {
            return false;
        }

        var sequence = new List<HtmlNode>();
        var lastIndex = index;

        while (lastIndex < nodes.Count && IsInlineSequenceCandidate(nodes[lastIndex]))
        {
            sequence.Add(nodes[lastIndex]);
            lastIndex++;
        }

        index = lastIndex - 1;
        if (sequence.Count == 0)
        {
            return true;
        }

        var sourceNode = TryGetInlineSequenceSourceNode(sequence);
        var text = ConvertNodesAsInline(sequence, context).ReplaceLineEndings("\n").Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        blocks = ConvertTextBlocks(text, quoteDepth, sourceNode, headingInference);
        return true;
    }

    private static List<string> ConvertTextBlocks(
        string text,
        int quoteDepth,
        HtmlNode? sourceNode = null,
        HeadingInference? headingInference = null)
    {
        var normalized = text.ReplaceLineEndings("\n").Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        if (sourceNode is not null &&
            (TryGetSemanticHeadingLevel(sourceNode, out var headingLevel) ||
             (headingInference is not null && headingInference.TryGetLevel(sourceNode, out headingLevel))))
        {
            var headingText = NormalizeHeadingText(normalized.Replace("\n", " ", StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(headingText))
            {
                return [ApplyQuotePrefix($"{new string('#', headingLevel)} {headingText}", quoteDepth)];
            }
        }

        var blocks = new List<string>();
        var paragraphLines = new List<string>();

        foreach (var rawLine in normalized.Split('\n'))
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraph(blocks, paragraphLines, quoteDepth);
                continue;
            }

            if (TryConvertTaskLine(line, out var taskLine))
            {
                FlushParagraph(blocks, paragraphLines, quoteDepth);
                blocks.Add(ApplyQuotePrefix(taskLine, quoteDepth));
                continue;
            }

            if (TryConvertLooseOrderedLine(line, out var orderedLine))
            {
                FlushParagraph(blocks, paragraphLines, quoteDepth);
                blocks.Add(ApplyQuotePrefix(orderedLine, quoteDepth));
                continue;
            }

            if (TryConvertLooseBulletLine(line, out var bulletLine))
            {
                FlushParagraph(blocks, paragraphLines, quoteDepth);
                blocks.Add(ApplyQuotePrefix(bulletLine, quoteDepth));
                continue;
            }

            if (IsHorizontalRuleText(line))
            {
                FlushParagraph(blocks, paragraphLines, quoteDepth);
                blocks.Add(ApplyQuotePrefix("---", quoteDepth));
                continue;
            }

            paragraphLines.Add(line);
        }

        FlushParagraph(blocks, paragraphLines, quoteDepth);
        return blocks;
    }

    private static void FlushParagraph(List<string> blocks, List<string> paragraphLines, int quoteDepth)
    {
        if (paragraphLines.Count == 0)
        {
            return;
        }

        blocks.Add(ApplyQuotePrefix(string.Join("\n", paragraphLines), quoteDepth));
        paragraphLines.Clear();
    }

    private static bool IsInlineSequenceCandidate(HtmlNode node)
    {
        if (ShouldSkip(node))
        {
            return false;
        }

        if (node.NodeType == HtmlNodeType.Text)
        {
            return !string.IsNullOrWhiteSpace(WebUtility.HtmlDecode(node.InnerText));
        }

        if (node.NodeType != HtmlNodeType.Element)
        {
            return false;
        }

        if (node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return node.Name is not ("ul" or "ol" or "table" or "blockquote" or "pre" or "hr") &&
            !IsBlockContainer(node) &&
            !IsStructuralContainer(node);
    }

    private static HtmlNode? TryGetInlineSequenceSourceNode(IReadOnlyList<HtmlNode> nodes)
    {
        var elements = nodes
            .Where(node => node.NodeType == HtmlNodeType.Element && !node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (elements.Count != 1)
        {
            return null;
        }

        var hasMeaningfulTextSibling = nodes.Any(node =>
            node.NodeType == HtmlNodeType.Text &&
            !string.IsNullOrWhiteSpace(WebUtility.HtmlDecode(node.InnerText)));

        return hasMeaningfulTextSibling ? null : elements[0];
    }

    private sealed class ConversionContext
    {
        private readonly byte[]? _fallbackImagePng;

        public ConversionContext(byte[]? fallbackImagePng, int _)
        {
            _fallbackImagePng = fallbackImagePng;
        }

        public bool CanUseFallbackImage => _fallbackImagePng is { Length: > 0 };

        public string UseFallbackImage()
        {
            return $"data:image/png;base64,{System.Convert.ToBase64String(_fallbackImagePng!)}";
        }
    }

    private sealed class HeadingInference
    {
        private readonly Dictionary<HtmlNode, int> _nodeLevels;

        private HeadingInference(Dictionary<HtmlNode, int> nodeLevels)
        {
            _nodeLevels = nodeLevels;
        }

        public static HeadingInference Analyze(HtmlNode root)
        {
            var hasSemanticHeading = root.DescendantsAndSelf().Any(node => TryGetSemanticHeadingLevel(node, out _));
            AppLogger.Debug($"Heading analysis: semantic heading present = {hasSemanticHeading}.");
            if (hasSemanticHeading)
            {
                AppLogger.Debug("Heading analysis: semantic heading detected, candidate heading inference disabled for the entire fragment.");
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var bodyFontSamples = CollectBodyFontSizes(root);
            if (bodyFontSamples.Count == 0)
            {
                AppLogger.Debug("Heading analysis: no font-size samples found for the current fragment.");
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var baseline = bodyFontSamples
                .GroupBy(size => size)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key)
                .First()
                .Key;
            AppLogger.Debug($"Heading analysis: body font baseline = {baseline:F2}pt.");

            var baseCandidates = CollectBaseCandidates(root);
            AppLogger.Debug($"Heading analysis: base candidate count = {baseCandidates.Count}.");
            if (baseCandidates.Count == 0)
            {
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var effectiveCandidates = baseCandidates
                .Where(candidate => candidate.FontSizePt >= baseline + 2d || candidate.FontSizePt >= baseline * 1.15d)
                .ToList();
            AppLogger.Debug($"Heading analysis: effective candidate count = {effectiveCandidates.Count}.");
            if (effectiveCandidates.Count == 0)
            {
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var orderedFontBands = effectiveCandidates
                .Select(candidate => candidate.FontSizePt)
                .Distinct()
                .OrderByDescending(size => size)
                .Take(3)
                .ToList();

            var mappingDescription = orderedFontBands.Count == 0
                ? "none"
                : string.Join(", ", orderedFontBands.Select((size, levelIndex) => $"{size:F2}pt=>H{levelIndex + 1}"));
            AppLogger.Debug($"Heading analysis: font bands = {mappingDescription}.");

            var nodeLevels = new Dictionary<HtmlNode, int>();
            foreach (var candidate in effectiveCandidates)
            {
                var bandIndex = orderedFontBands.FindIndex(size => Math.Abs(size - candidate.FontSizePt) < 0.01d);
                if (bandIndex >= 0)
                {
                    nodeLevels[candidate.Node] = bandIndex + 1;
                }
            }

            return new HeadingInference(nodeLevels);
        }

        public bool TryGetLevel(HtmlNode node, out int level) => _nodeLevels.TryGetValue(node, out level);

        private static List<double> CollectBodyFontSizes(HtmlNode root)
        {
            var sizes = new List<double>();

            foreach (var node in root.DescendantsAndSelf())
            {
                if (ShouldSkip(node) || node.NodeType != HtmlNodeType.Element)
                {
                    continue;
                }

                var text = ExtractCandidateText(node);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (TryExtractFontSize(node, out var fontSizePt))
                {
                    sizes.Add(Math.Round(fontSizePt, 2));
                }
            }

            return sizes;
        }

        private static List<HeadingCandidate> CollectBaseCandidates(HtmlNode root)
        {
            var candidates = new List<HeadingCandidate>();

            foreach (var node in root.Descendants())
            {
                if (!IsCandidateBlock(node))
                {
                    continue;
                }

                if (HasBlockedAncestor(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is inside a blocked container.");
                    continue;
                }

                if (IsStructuralContainer(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a structural container.");
                    continue;
                }

                if (node.SelectSingleNode(".//br") is not null)
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is not a single visual line.");
                    continue;
                }

                if (TryConvertOfficeListParagraph(node, 0, out _))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is an Office list paragraph.");
                    continue;
                }

                if (HasQuoteStyle(node) || IsCodeBlock(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a quote or code block.");
                    continue;
                }

                var text = ExtractCandidateText(node);
                if (string.IsNullOrWhiteSpace(text))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' text is empty.");
                    continue;
                }

                if (text.Length > MaximumCandidateHeadingLength)
                {
                    AppLogger.Debug($"Heading candidate rejected: '{text}' exceeds {MaximumCandidateHeadingLength} characters.");
                    continue;
                }

                if (IsDisallowedHeadingTerminator(text[^1]))
                {
                    AppLogger.Debug($"Heading candidate rejected: '{text}' ends with disallowed punctuation.");
                    continue;
                }

                if (!TryExtractFontSize(node, out var fontSizePt))
                {
                    AppLogger.Debug($"Heading candidate rejected: '{text}' has no extractable font-size.");
                    continue;
                }

                candidates.Add(new HeadingCandidate(node, text, Math.Round(fontSizePt, 2)));
            }

            return candidates;
        }
    }

    private sealed record HeadingCandidate(HtmlNode Node, string Text, double FontSizePt);
}
