using Html2Markdown;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class MarkdownLineSyntaxTests
{
    [Fact]
    public void TryParseSourceTask_ParsesCheckedStateAndContent()
    {
        Assert.True(MarkdownLineSyntax.TryParseSourceTask("☑ finish docs", out var content, out var isChecked));
        Assert.Equal("finish docs", content);
        Assert.True(isChecked);
        Assert.Equal("- [x] finish docs", MarkdownLineSyntax.RenderMarkdownTask(content, isChecked));
    }

    [Fact]
    public void TryParseSourceOrderedAndBulletListItems_PreserveMarkdownRendering()
    {
        Assert.True(MarkdownLineSyntax.TryParseSourceOrderedListItem("2. ship it", out var marker, out var orderedContent));
        Assert.Equal("2.", marker);
        Assert.Equal("ship it", orderedContent);
        Assert.Equal("2. ship it", MarkdownLineSyntax.RenderMarkdownOrderedListItem(marker, orderedContent));

        Assert.True(MarkdownLineSyntax.TryParseSourceBulletListItem("• keep it small", out var bulletContent));
        Assert.Equal("keep it small", bulletContent);
        Assert.Equal("- keep it small", MarkdownLineSyntax.RenderMarkdownBulletListItem(bulletContent));
    }

    [Fact]
    public void StripPrefixes_RemovesSharedMarkdownAndSourcePrefixes()
    {
        Assert.Equal("done", MarkdownLineSyntax.StripMarkdownPrefix("> ## 1. [x] done"));
        Assert.Equal("done", MarkdownLineSyntax.StripSourcePrefix("1. ☑ done"));
    }
}
