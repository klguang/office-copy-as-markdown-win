using HtmlAgilityPack;

namespace Html2Markdown;

public sealed class StandardHtmlDialectAdapter : IHtmlDialectAdapter
{
    public static StandardHtmlDialectAdapter Instance { get; } = new();

    private StandardHtmlDialectAdapter()
    {
    }

    public bool TryGetSemanticHeadingLevel(HtmlNode node, out int level)
    {
        level = 0;
        return false;
    }

    public bool IsQuoteBlock(HtmlNode node) => false;

    public bool TryConvertListLikeBlock(HtmlNode node, int quoteDepth, out string markdown)
    {
        markdown = string.Empty;
        return false;
    }

    public bool ShouldBlockHeadingInference(HtmlNode node) => false;
}
