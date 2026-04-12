using Html2Markdown;
using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class HtmlToMarkdownApiTests
{
    [Fact]
    public void NormalizeHtml_LeavesInputUntouched()
    {
        const string html = "<html><body><p>kept</p></body></html>";

        var normalized = HtmlToMarkdownPipeline.NormalizeHtml(html);

        Assert.Equal(html, normalized);
    }

    [Fact]
    public void ExtractFragment_ReturnsClipboardFragment()
    {
        const string fragment = "<p>hello <strong>world</strong></p>";
        var cfHtml = BuildCfHtml(fragment);

        var extracted = HtmlToMarkdownPipeline.ExtractFragment(cfHtml);

        Assert.Equal(fragment, extracted);
    }

    [Fact]
    public void Convert_HandlesRawCfHtml()
    {
        const string fragment = "<html><body><p>Hello</p><ul><li>world</li></ul></body></html>";
        var markdown = HtmlToMarkdownConverter.Convert(
            BuildCfHtml(fragment),
            new HtmlToMarkdownOptions
            {
                ExtractClipboardFragment = true
            });

        Assert.Contains("Hello", markdown);
        Assert.Contains("- world", markdown);
    }

    [Fact]
    public void Convert_HandlesNormalHtmlBody()
    {
        const string html = """
            <html>
            <body>
              <h2>Title</h2>
              <p>Paragraph</p>
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownConverter.Convert(html);

        Assert.Contains("## Title", markdown);
        Assert.Contains("Paragraph", markdown);
    }

    [Fact]
    public void Convert_UsesFallbackImageWhenImageSourceIsMissing()
    {
        const string html = """
            <html>
            <body>
              <img alt="clipboard image" />
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownConverter.Convert(
            html,
            new HtmlToMarkdownOptions
            {
                FallbackImagePng = [1, 2, 3, 4]
            });

        Assert.StartsWith("![clipboard image](data:image/png;base64,", markdown);
    }

    [Fact]
    public void Convert_FallsBackToConservativeMarkdownWhenHtmlProducesNoContent()
    {
        const string html = "<html><body></body></html>";
        const string sourceText = """
            Heading
            • second line
            """;

        var markdown = HtmlToMarkdownConverter.Convert(
            html,
            new HtmlToMarkdownOptions
            {
                SourceText = sourceText,
                RepairMode = HtmlToMarkdownRepairMode.IfSourceTextAvailable
            });

        Assert.Contains("Heading", markdown);
        Assert.Contains("- second line", markdown);
    }

    [Fact]
    public void RepairMarkdown_ExposesRepairResultThroughPublicPipeline()
    {
        const string markdown = "# Heading";
        const string sourceText = """
            Heading
            • missing line
            """;

        var repaired = HtmlToMarkdownPipeline.RepairMarkdown(markdown, sourceText);

        Assert.True(repaired.IsComplete);
        Assert.Contains("- missing line", repaired.Markdown);
        Assert.Empty(repaired.MissingLines);
    }

    [Fact]
    public void Convert_DoesNotTreatOfficeHeadingMarkupAsSemanticHeadingWithoutDialectAdapter()
    {
        const string html = """
            <html>
            <body>
              <div><span style="font-size:18pt">大字号候选</span></div>
              <p style="mso-style-name:Heading 2">正式标题</p>
              <p><span style="font-size:12pt">正文内容</span></p>
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownConverter.Convert(html);

        Assert.DoesNotContain("## 正式标题", markdown);
        Assert.Contains("大字号候选", markdown);
        Assert.Contains("正式标题", markdown);
    }

    [Fact]
    public void Convert_DoesNotTreatOfficeListParagraphAsListWithoutDialectAdapter()
    {
        const string html = """
            <html>
            <body>
              <p class="MsoListParagraph" style="mso-list:l0 level1 lfo1">First</p>
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownConverter.Convert(html);

        Assert.DoesNotContain("- First", markdown);
        Assert.Contains("First", markdown);
    }

    [Fact]
    public void Convert_AllowsCustomHeadingInferenceStrategyToOverrideDefaultBehavior()
    {
        const string html = """
            <html>
            <body>
              <p>Custom heading</p>
              <p>Body</p>
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownConverter.Convert(
            html,
            new HtmlToMarkdownOptions
            {
                HeadingInferenceStrategy = new ForceFirstParagraphHeadingStrategy()
            });

        Assert.Contains("# Custom heading", markdown);
        Assert.Contains("Body", markdown);
    }

    [Fact]
    public void ConvertFragment_AllowsCustomDialectAdapterToOverrideSemanticHeadingRecognition()
    {
        const string html = """
            <html>
            <body>
              <div data-heading-level="3">Adapter heading</div>
              <p>Body</p>
            </body>
            </html>
            """;

        var markdown = HtmlToMarkdownPipeline.ConvertFragment(
            html,
            new HtmlToMarkdownOptions
            {
                DialectAdapter = new DataHeadingDialectAdapter()
            });

        Assert.Contains("### Adapter heading", markdown);
        Assert.Contains("Body", markdown);
    }

    private static string BuildCfHtml(string fragment)
    {
        const string headerTemplate = "Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n";
        const string startMarker = "<!--StartFragment-->";
        const string endMarker = "<!--EndFragment-->";
        const string prefix = "<html><body>";
        const string suffix = "</body></html>";

        var htmlBody = $"{prefix}{startMarker}{fragment}{endMarker}{suffix}";
        var startHtml = headerTemplate.Length;
        var startFragment = startHtml + prefix.Length + startMarker.Length;
        var endFragment = startFragment + fragment.Length;
        var endHtml = startHtml + htmlBody.Length;

        var header = $"""
            Version:0.9
            StartHTML:{startHtml:D10}
            EndHTML:{endHtml:D10}
            StartFragment:{startFragment:D10}
            EndFragment:{endFragment:D10}
            """.ReplaceLineEndings("\r\n") + "\r\n";

        return header + htmlBody;
    }

    private sealed class ForceFirstParagraphHeadingStrategy : IHeadingInferenceStrategy
    {
        public IReadOnlyDictionary<HtmlNode, int> InferHeadingLevels(HtmlNode root, IHtmlDialectAdapter dialectAdapter)
        {
            var firstParagraph = root.Descendants("p").First();
            return new Dictionary<HtmlNode, int>
            {
                [firstParagraph] = 1
            };
        }
    }

    private sealed class DataHeadingDialectAdapter : IHtmlDialectAdapter
    {
        public bool TryGetSemanticHeadingLevel(HtmlNode node, out int level)
        {
            if (int.TryParse(node.GetAttributeValue("data-heading-level", string.Empty), out level))
            {
                return true;
            }

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
}
