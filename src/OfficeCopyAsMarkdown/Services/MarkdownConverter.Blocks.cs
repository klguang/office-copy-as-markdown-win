using System.Net;
using System.Text;
using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Services;

internal static partial class MarkdownConverter
{
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
                if (TryConvertSingleCellTableAsQuote(node, context, headingInference, quoteDepth, out var quotedTable))
                {
                    blocks.Add(quotedTable);
                    continue;
                }

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

    private static bool TryConvertSingleCellTableAsQuote(
        HtmlNode tableNode,
        ConversionContext context,
        HeadingInference headingInference,
        int quoteDepth,
        out string markdown)
    {
        markdown = string.Empty;

        var rows = tableNode.SelectNodes(".//tr");
        if (rows is null || rows.Count != 1)
        {
            return false;
        }

        var cells = rows[0].Elements("th").Concat(rows[0].Elements("td")).ToList();
        if (cells.Count != 1)
        {
            return false;
        }

        var cell = cells[0];
        var blocks = ConvertBlocks(cell.ChildNodes, context, headingInference, quoteDepth + 1);
        if (blocks.Count == 0)
        {
            var fallback = ConvertTextBlocks(ConvertInlines(cell, context), quoteDepth + 1, cell, headingInference);
            blocks.AddRange(fallback);
        }

        markdown = string.Join("\n\n", blocks.Where(block => !string.IsNullOrWhiteSpace(block))).Trim();
        return !string.IsNullOrWhiteSpace(markdown);
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
}
