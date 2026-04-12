using OfficeCopyAsMarkdown.Services;
using System.Text.RegularExpressions;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class MarkdownConverterTests
{
    [Fact]
    public void Convert_PreservesInlineSequencesInsideStructuralContainers()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "OneNoteUserJourney.html");
        var html = File.ReadAllText(fixturePath, System.Text.Encoding.UTF8);

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        AssertUserJourneyMarkdown(markdown);
    }

    [Fact]
    public void Convert_PreservesProgramLogHtmlFixture()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "OneNoteUserJourneyFromLog.html");
        var html = File.ReadAllText(fixturePath, System.Text.Encoding.UTF8);

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        AssertUserJourneyMarkdown(markdown);
    }

    [Fact]
    public void Convert_PreservesStandardUnorderedLists()
    {
        const string html = """
            <html>
            <body>
              <ul>
                <li>first item</li>
                <li>second item</li>
              </ul>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Equal("- first item\n- second item", markdown);
    }

    [Fact]
    public void Convert_PreservesNonStandardSiblingNestedLists()
    {
        const string html = """
            <html>
            <body>
              <ol>
                <li><b>发现</b>：</li>
                <ul>
                  <li>浏览合集库 -&gt; 查看合集信息（封面、简介、是否公开/免费）</li>
                </ul>
                <li><b>获取</b>：</li>
                <ul>
                  <li>免费合集 -&gt; 收藏或添加到学习计划</li>
                  <li>付费合集 -&gt; 收藏（购买逻辑暂不实现）</li>
                </ul>
              </ol>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("1. **发现**：", markdown);
        Assert.Contains("- 浏览合集库 -> 查看合集信息（封面、简介、是否公开/免费）", markdown);
        Assert.Contains("2. **获取**：", markdown);
        Assert.Contains("- 免费合集 -> 收藏或添加到学习计划", markdown);
        Assert.Contains("- 付费合集 -> 收藏（购买逻辑暂不实现）", markdown);
    }

    [Fact]
    public void Convert_WhenSemanticHeadingExists_DisablesCandidateHeadingInference()
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

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 正式标题", markdown);
        Assert.Contains("大字号候选", markdown);
        Assert.DoesNotContain("# 大字号候选", markdown);
        Assert.DoesNotContain("## 大字号候选", markdown);
        Assert.DoesNotContain("### 大字号候选", markdown);
        Assert.DoesNotContain("#### 大字号候选", markdown);
    }

    [Fact]
    public void Convert_InfersCandidateHeadingsByFontBands()
    {
        const string html = """
            <html>
            <body>
              <div><span style="font-size:20pt;font-weight:bold;">主标题</span></div>
              <div><span style="font-size:18pt;font-weight:bold;">次标题</span></div>
              <div><span style="font-size:16pt;font-weight:bold;">三级标题</span></div>
              <div><span style="font-size:14pt;font-weight:bold;">四级标题</span></div>
              <p><span style="font-size:12pt;">正文内容</span></p>
              <p><span style="font-size:12pt;">第二段正文</span></p>
              <p><span style="font-size:12pt;">第三段正文</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("# 主标题", markdown);
        Assert.Contains("## 次标题", markdown);
        Assert.Contains("### 三级标题", markdown);
        Assert.Contains("#### 四级标题", markdown);
        Assert.Contains("正文内容", markdown);
    }

    [Fact]
    public void Convert_InfersLongMainTitleWithinUpdatedLengthLimit()
    {
        const string html = """
            <html>
            <body>
              <p><span style="font-size:20pt;font-weight:bold;">摸鱼卡片（manycards.cn）产品规格</span></p>
              <p><span style="font-size:18pt;font-weight:bold;">核心概念</span></p>
              <p><span style="font-size:16pt;font-weight:bold;">2.1 学习者路径</span></p>
              <p><span style="font-size:12pt;">正文内容</span></p>
              <p><span style="font-size:12pt;">第二段正文</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 摸鱼卡片（manycards.cn）产品规格", markdown);
        Assert.Contains("### 核心概念", markdown);
        Assert.Contains("#### 2.1 学习者路径", markdown);
    }

    [Fact]
    public void Convert_InfersBoldCandidateHeadingsAtBodyBaseline()
    {
        const string html = """
            <html>
            <body>
              <div><span style="font-size:18pt;font-weight:bold;">一级标题</span></div>
              <div><span style="font-size:16pt;font-weight:bold;">二级标题</span></div>
              <div><span style="font-size:14pt;font-weight:bold;">三级标题</span></div>
              <div><span style="font-size:12pt;font-weight:bold;">四级标题</span></div>
              <p><span style="font-size:12pt;">正文内容</span></p>
              <p><span style="font-size:12pt;">第二段正文</span></p>
              <div><span style="font-size:12pt;">普通短句</span></div>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("# 一级标题", markdown);
        Assert.Contains("## 二级标题", markdown);
        Assert.Contains("### 三级标题", markdown);
        Assert.Contains("#### 四级标题", markdown);
        Assert.Contains("普通短句", markdown);
        Assert.DoesNotContain("# 普通短句", markdown);
        Assert.DoesNotContain("## 普通短句", markdown);
        Assert.DoesNotContain("### 普通短句", markdown);
        Assert.DoesNotContain("#### 普通短句", markdown);
    }

    [Fact]
    public void Convert_DoesNotInferNonBoldOrSmallerBoldCandidatesAtBodyBaseline()
    {
        const string html = """
            <html>
            <body>
              <div><span style="font-size:14pt;font-weight:bold;">有效标题</span></div>
              <div><span style="font-size:12pt;">同字号短句</span></div>
              <div><span style="font-size:11pt;font-weight:bold;">更小粗体</span></div>
              <p><span style="font-size:12pt;">正文内容</span></p>
              <p><span style="font-size:12pt;">第二段正文</span></p>
              <p><span style="font-size:12pt;">第三段正文</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 有效标题", markdown);
        Assert.Contains("同字号短句", markdown);
        Assert.Contains("更小粗体", markdown);
        Assert.DoesNotContain("# 同字号短句", markdown);
        Assert.DoesNotContain("## 同字号短句", markdown);
        Assert.DoesNotContain("### 同字号短句", markdown);
        Assert.DoesNotContain("#### 同字号短句", markdown);
        Assert.DoesNotContain("# 更小粗体", markdown);
        Assert.DoesNotContain("## 更小粗体", markdown);
        Assert.DoesNotContain("### 更小粗体", markdown);
        Assert.DoesNotContain("#### 更小粗体", markdown);
    }

    [Fact]
    public void Convert_FollowsHeadingPunctuationRulesFromDocument()
    {
        const string html = """
            <html>
            <body>
              <div><span style="font-size:16pt;font-weight:bold;">允许：</span></div>
              <div><span style="font-size:16pt;font-weight:bold;">不允许。</span></div>
              <p><span style="font-size:12pt;">正文内容</span></p>
              <p><span style="font-size:12pt;">第二段正文</span></p>
              <p><span style="font-size:12pt;">第三段正文</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 允许：", markdown);
        Assert.Contains("不允许。", markdown);
        Assert.DoesNotContain("# 不允许。", markdown);
    }

    [Fact]
    public void Convert_MergesAdjacentBoldSegmentsWithoutExtraMarkers()
    {
        const string html = """
            <html>
            <body>
              <p>
                <span style="font-weight:bold;">摸鱼卡片（</span>
                <span style="font-weight:bold;">manycards.cn</span>
                <span style="font-weight:bold;">）产品规格</span>
              </p>
              <table>
                <tr>
                  <th>名称</th>
                  <th>描述</th>
                </tr>
                <tr>
                  <td>
                    <span style="font-weight:bold;">收藏</span>
                    <span style="font-weight:bold;">Favorite</span>
                  </td>
                  <td>收藏合集或章节，方便后续学习或试学</td>
                </tr>
                <tr>
                  <td>
                    <span style="font-weight:bold;">发布</span>
                    <span style="font-weight:bold;">/</span>
                    <span style="font-weight:bold;">分享</span>
                  </td>
                  <td>设置合集可见性（私有/免费）</td>
                </tr>
              </table>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.DoesNotContain("****", markdown);
        Assert.Contains("**摸鱼卡片（manycards.cn）产品规格**", markdown);
        Assert.Contains("| **收藏 Favorite** | 收藏合集或章节，方便后续学习或试学 |", markdown);
        Assert.Contains("| **发布/分享** | 设置合集可见性（私有/免费） |", markdown);
    }

    [Fact]
    public void Convert_RendersSingleCellTableAsBlockquote()
    {
        const string html = """
            <html>
            <body>
              <table>
                <tr>
                  <td>
                    <ul>
                      <li><b>Entity：</b>持久化对象，对应数据库中的一条记录</li>
                      <li><b>VO（Value Object）：</b>值对象，即调用service参数或者返回值</li>
                    </ul>
                    <p><b>重要：</b>如果VO对象满足Controller返回结果，不用新建Response对象；java 值转换必须使用mapstruct</p>
                  </td>
                </tr>
              </table>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.DoesNotContain("| --- |", markdown);
        Assert.DoesNotContain("| **Entity：**", markdown);
        Assert.Contains("> - **Entity：** 持久化对象，对应数据库中的一条记录", markdown);
        Assert.Contains("> - **VO（Value Object）：** 值对象，即调用service参数或者返回值", markdown);
        Assert.Contains("> **重要：** 如果VO对象满足Controller返回结果，不用新建Response对象；java 值转换必须使用mapstruct", markdown);
    }

    [Fact]
    public void Convert_InsertsSpaceAfterBoldLabelBeforePlainText()
    {
        const string html = """
            <html>
            <body>
              <p><b>Entity：</b>持久化对象，对应数据库中的一条记录</p>
              <p><b>重要：</b>如果VO对象满足Controller返回结果，不用新建Response对象</p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("**Entity：** 持久化对象，对应数据库中的一条记录", markdown);
        Assert.Contains("**重要：** 如果VO对象满足Controller返回结果，不用新建Response对象", markdown);
    }

    [Fact]
    public void Convert_InsertsSpaceBeforeFollowingBoldLabelsAndAroundPlus()
    {
        const string html = """
            <html>
            <body>
              <p>
                <span style="font-weight:bold;">阶段</span><span>：</span><span>MVP</span>
                <span style="font-weight:bold;">目标用户</span><span>：</span><span>UGC创作者</span><span>+</span><span>碎片化自学学习者</span>
                <span style="font-weight:bold;">核心价值</span><span>：</span><span>支持用户创建官方模板卡片</span>
              </p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("**阶段：** MVP **目标用户：** UGC创作者 + 碎片化自学学习者 **核心价值：** 支持用户创建官方模板卡片", markdown);
    }

    [Fact]
    public void Convert_InfersHeadingWhenTextIsMostlyBold()
    {
        const string html = """
            <html>
            <body>
              <p><span style="font-size:12pt;font-weight:bold;">验收标准（</span><span style="font-size:12pt;">EARS</span><span style="font-size:12pt;font-weight:bold;">语法）</span></p>
              <p><span style="font-size:12pt;">- WHEN 用户提交正确的注册信息 THEN 系统创建账号、发送验证邮件、返回JWT token。</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 验收标准（EARS 语法）", markdown);
    }

    [Fact]
    public void Convert_StartsInferredHeadingsAtLevelTwoWhenThereAreFewerThanFourBands()
    {
        const string html = """
            <html>
            <body>
              <p><span style="font-size:12pt;font-weight:bold;">验收标准（</span><span style="font-size:12pt;font-weight:bold;">EARS</span><span style="font-size:12pt;font-weight:bold;">语法）</span></p>
              <p><span style="font-size:12pt;">- WHEN 用户提交正确的注册信息 THEN 系统创建账号、发送验证邮件、返回JWT token。</span></p>
              <p><span style="font-size:12pt;font-weight:bold;">非功能约束</span></p>
              <p><span style="font-size:12pt;">每个模块可以列出与功能相关的非功能需求：</span></p>
              <p><span style="font-size:12pt;font-weight:bold;">附加说明</span></p>
              <p><span style="font-size:12pt;">业务规则：如“同一用户不能重复创建同名合集”</span></p>
            </body>
            </html>
            """;

        var markdown = MarkdownConverter.Convert(html, fallbackImagePng: null);

        Assert.Contains("## 验收标准（EARS语法）", markdown);
        Assert.Contains("## 非功能约束", markdown);
        Assert.Contains("## 附加说明", markdown);
        Assert.DoesNotContain("\n# 验收标准（EARS语法）\n", $"\n{markdown}\n");
        Assert.DoesNotContain("\n# 非功能约束\n", $"\n{markdown}\n");
        Assert.DoesNotContain("\n# 附加说明\n", $"\n{markdown}\n");
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
