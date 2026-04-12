using Html2Markdown;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class MarkdownContentGuardTests
{
    [Fact]
    public void ShouldKeepMarkdown_ReturnsFalseWhenLargeSectionsDisappear()
    {
        const string sourceText = """
            用户路径

            同一用户可以是学习者，也可以是创作者。

            2.1 学习者路径
            1. 发现：
            浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）
            2. 获取：
            免费合集 -> 收藏或添加到学习计划
            付费合集 -> 收藏（购买逻辑暂不实现）
            """;

        const string markdown = """
            # 用户路径

            ## 2.1 学习者路径

            1. **发现**：
            2. **获取**：
            """;

        Assert.False(HtmlToMarkdownPipeline.ShouldKeepMarkdown(markdown, sourceText));
    }

    [Fact]
    public void ShouldKeepMarkdown_ReturnsTrueForEquivalentVisibleContent()
    {
        const string sourceText = """
            用户路径
            同一用户可以是学习者，也可以是创作者。
            1. 发现：
            浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）
            """;

        const string markdown = """
            # 用户路径

            同一用户可以是学习者，也可以是创作者。

            1. **发现**：
            - 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）
            """;

        Assert.True(HtmlToMarkdownPipeline.ShouldKeepMarkdown(markdown, sourceText));
    }

    [Fact]
    public void RepairMarkdown_ReinsertsMissingLinesWithoutCreatingNewHeadings()
    {
        const string sourceText = """
            用户路径
            1. 发现：
            ○ 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）
            2. 获取：
            ○ 免费合集 -> 收藏或添加到学习计划
            """;

        const string markdown = """
            # 用户路径

            1. **发现**：
            """;

        var repaired = HtmlToMarkdownPipeline.RepairMarkdown(markdown, sourceText);

        Assert.True(repaired.IsComplete);
        Assert.Contains("- 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）", repaired.Markdown);
        Assert.Contains("2. 获取：", repaired.Markdown);
        Assert.DoesNotContain("# 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）", repaired.Markdown);
        Assert.Empty(repaired.MissingLines);
    }

    [Fact]
    public void BuildConservativeMarkdown_PreservesMarkdownWithoutInferringHeadings()
    {
        const string sourceText = """
            用户路径
            2.1 学习者路径
            ○ 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）
            """;

        var markdown = HtmlToMarkdownPipeline.BuildConservativeMarkdown(sourceText);

        Assert.Contains("用户路径", markdown);
        Assert.Contains("2.1 学习者路径", markdown);
        Assert.Contains("- 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）", markdown);
        Assert.DoesNotContain("# 2.1 学习者路径", markdown);
    }

    [Fact]
    public void RepairMarkdown_DoesNotDuplicateMarkdownTableContentWhenSourceUsesTabs()
    {
        const string sourceText = """
            名称	描述
            Member	平台统一用户角色，既可以创建卡片/合集（创作者），也可以学习/收藏（学习者）
            """;

        const string markdown = """
            | 名称 | 描述 |
            | --- | --- |
            | Member | 平台统一用户角色，既可以创建卡片/合集（创作者），也可以学习/收藏（学习者） |
            """;

        Assert.True(HtmlToMarkdownPipeline.ShouldKeepMarkdown(markdown, sourceText));

        var repaired = HtmlToMarkdownPipeline.RepairMarkdown(markdown, sourceText);

        Assert.True(repaired.IsComplete);
        Assert.Equal(markdown, repaired.Markdown);
        Assert.Empty(repaired.MissingLines);
    }

    [Fact]
    public void RepairMarkdown_DoesNotDuplicateMixedChineseEnglishSummaryLine()
    {
        const string sourceText = """
            阶段：MVP 目标用户：UGC创作者 + 碎片化自学学习者 核心价值：支持用户创建官方模板卡片、组织合集与章节、制定学习计划、个性化笔记及收藏
            """;

        const string markdown = """
            **阶段：** MVP **目标用户：** UGC创作者 + 碎片化自学学习者 **核心价值：** 支持用户创建官方模板卡片、组织合集与章节、制定学习计划、个性化笔记及收藏
            """;

        Assert.True(HtmlToMarkdownPipeline.ShouldKeepMarkdown(markdown, sourceText));

        var repaired = HtmlToMarkdownPipeline.RepairMarkdown(markdown, sourceText);

        Assert.True(repaired.IsComplete);
        Assert.Equal(markdown, repaired.Markdown);
        Assert.Empty(repaired.MissingLines);
    }
}
