using HtmlAgilityPack;

namespace Html2Markdown;

public interface IHtmlDialectAdapter
{
    bool TryGetSemanticHeadingLevel(HtmlNode node, out int level);

    bool IsQuoteBlock(HtmlNode node);

    bool TryConvertListLikeBlock(HtmlNode node, int quoteDepth, out string markdown);

    bool ShouldBlockHeadingInference(HtmlNode node);
}
