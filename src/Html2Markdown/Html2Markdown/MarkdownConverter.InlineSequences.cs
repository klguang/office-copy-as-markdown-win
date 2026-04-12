using System.Net;
using HtmlAgilityPack;

namespace Html2Markdown;

internal static partial class MarkdownConverter
{
    private static bool TryConvertInlineSequence(
        HtmlNodeCollection nodes,
        ref int index,
        ConversionContext context,
        IReadOnlyDictionary<HtmlNode, int> inferredHeadingLevels,
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

        blocks = ConvertTextBlocks(text, quoteDepth, context, inferredHeadingLevels, sourceNode);
        return true;
    }

    private static List<string> ConvertTextBlocks(
        string text,
        int quoteDepth,
        ConversionContext context,
        IReadOnlyDictionary<HtmlNode, int> inferredHeadingLevels,
        HtmlNode? sourceNode = null)
    {
        var normalized = text.ReplaceLineEndings("\n").Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        if (sourceNode is not null &&
            (TryGetSemanticHeadingLevel(sourceNode, context.DialectAdapter, out var headingLevel) ||
             inferredHeadingLevels.TryGetValue(sourceNode, out headingLevel)))
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

        public ConversionContext(byte[]? fallbackImagePng, IHtmlDialectAdapter dialectAdapter)
        {
            _fallbackImagePng = fallbackImagePng;
            DialectAdapter = dialectAdapter;
        }

        public IHtmlDialectAdapter DialectAdapter { get; }

        public bool CanUseFallbackImage => _fallbackImagePng is { Length: > 0 };

        public string UseFallbackImage()
        {
            return $"data:image/png;base64,{System.Convert.ToBase64String(_fallbackImagePng!)}";
        }
    }
}
