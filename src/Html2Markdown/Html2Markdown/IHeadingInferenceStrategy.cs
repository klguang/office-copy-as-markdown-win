using HtmlAgilityPack;

namespace Html2Markdown;

public interface IHeadingInferenceStrategy
{
    IReadOnlyDictionary<HtmlNode, int> InferHeadingLevels(HtmlNode root, IHtmlDialectAdapter dialectAdapter);
}
