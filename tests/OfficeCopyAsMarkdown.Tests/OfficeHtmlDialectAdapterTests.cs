using System.Text.RegularExpressions;
using Html2Markdown;
using OfficeCopyAsMarkdown.Services;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class OfficeHtmlDialectAdapterTests
{
    [Fact]
    public void Convert_WhenOfficeSemanticHeadingExists_DisablesCandidateHeadingInference()
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

        var markdown = HtmlToMarkdownPipeline.ConvertFragment(
            html,
            new HtmlToMarkdownOptions
            {
                DialectAdapter = OfficeHtmlDialectAdapter.Instance
            });

        Assert.Contains("## 正式标题", markdown);
        Assert.Contains("大字号候选", markdown);
        Assert.DoesNotContain("# 大字号候选", markdown);
        Assert.DoesNotContain("## 大字号候选", markdown);
        Assert.DoesNotContain("### 大字号候选", markdown);
        Assert.DoesNotContain("#### 大字号候选", markdown);
    }

    [Fact]
    public void Convert_PreservesInlineSequencesInsideStructuralContainersForOfficeFixtures()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "OneNoteUserJourney.html");
        var html = OfficeHtmlNormalizer.Normalize(File.ReadAllText(fixturePath));

        var markdown = HtmlToMarkdownPipeline.ConvertFragment(
            html,
            new HtmlToMarkdownOptions
            {
                DialectAdapter = OfficeHtmlDialectAdapter.Instance
            });

        AssertUserJourneyMarkdown(markdown);
    }

    [Fact]
    public void Convert_PreservesProgramLogHtmlFixtureForOfficeDialect()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "OneNoteUserJourneyFromLog.html");
        var html = OfficeHtmlNormalizer.Normalize(File.ReadAllText(fixturePath));

        var markdown = HtmlToMarkdownPipeline.ConvertFragment(
            html,
            new HtmlToMarkdownOptions
            {
                DialectAdapter = OfficeHtmlDialectAdapter.Instance
            });

        AssertUserJourneyMarkdown(markdown);
    }

    private static void AssertUserJourneyMarkdown(string markdown)
    {
        AssertContainsNormalized(markdown, "# 用户路径");
        AssertContainsNormalized(markdown, "同一用户可以是学习者，也可以是创作者。");
        AssertContainsNormalized(markdown, "## 2.1 学习者路径");
        AssertContainsNormalized(markdown, "1. **发现**：");
        AssertContainsNormalized(markdown, "- 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）");
        AssertContainsNormalized(markdown, "2. **获取**：");
        AssertContainsNormalized(markdown, "- 免费合集 -> 收藏或添加到学习计划");
        AssertContainsNormalized(markdown, "- 付费合集 -> 收藏（购买逻辑暂不实现）");
        AssertContainsNormalized(markdown, "3. **个性化**：");
        AssertContainsNormalized(markdown, "- 对卡片添加笔记（文本/手写）");
        AssertContainsNormalized(markdown, "- 跟踪完成情况（打卡）");
        AssertContainsNormalized(markdown, "4. **坚持**：");
        AssertContainsNormalized(markdown, "- 根据学习计划完成每日任务");
        AssertContainsNormalized(markdown, "- 系统记录完成量和用时");
        AssertContainsNormalized(markdown, "5. **分享**：");
        AssertContainsNormalized(markdown, "- 分享合集、章节或卡片链接");
        AssertContainsNormalized(markdown, "- 链接附带来源信息，受访问权限控制");
        AssertContainsNormalized(markdown, "## 2.2 创作者路径");
        AssertContainsNormalized(markdown, "1. **创建合集**：");
        AssertContainsNormalized(markdown, "- 新建合集 -> 可选创建章节（最多一层子合集）");
        AssertContainsNormalized(markdown, "- 添加卡片到合集（从别的合集添加），或者创建卡片 -> 支持多种官方模板卡片混合");
        AssertContainsNormalized(markdown, "2. **创建卡片**：");
        AssertContainsNormalized(markdown, "- 选择官方模板 -> 填写内容 -> 保存卡片");
        AssertContainsNormalized(markdown, "3. **管理卡片**：");
        AssertContainsNormalized(markdown, "- 同一张卡片只能属于一个合集或章节");
        AssertContainsNormalized(markdown, "- 支持排序及编辑内容");
        AssertContainsNormalized(markdown, "- 付费合集可部分可见");
        AssertContainsNormalized(markdown, "4. **发布/分享**：");
        AssertContainsNormalized(markdown, "- 设置合集可见性（私有/免费）");
        AssertContainsNormalized(markdown, "- 分享合集或章节链接");
    }

    private static void AssertContainsNormalized(string actual, string expected)
    {
        Assert.Contains(NormalizeForAssertion(expected), NormalizeForAssertion(actual));
    }

    private static string NormalizeForAssertion(string markdown)
    {
        var normalized = markdown.ReplaceLineEndings("\n")
            .Replace("→", "->", StringComparison.Ordinal)
            .Replace("、", "，", StringComparison.Ordinal);

        normalized = normalized
            .Replace("**", string.Empty, StringComparison.Ordinal)
            .Replace("__", string.Empty, StringComparison.Ordinal)
            .Replace("~~", string.Empty, StringComparison.Ordinal)
            .Replace("`", string.Empty, StringComparison.Ordinal);

        normalized = Regex.Replace(normalized, @"\s*/\s*", "/");
        normalized = Regex.Replace(normalized, @"\s*：", "：");
        normalized = Regex.Replace(normalized, @"\s*->\s*", " -> ");
        normalized = Regex.Replace(normalized, @"[ \t]+", " ");
        normalized = Regex.Replace(normalized, @"\n{2,}", "\n\n");

        return normalized.Trim();
    }
}
