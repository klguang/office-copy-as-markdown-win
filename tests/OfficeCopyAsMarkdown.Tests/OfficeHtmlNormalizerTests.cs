using OfficeCopyAsMarkdown.Services;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class OfficeHtmlNormalizerTests
{
    [Fact]
    public void Normalize_StripsOfficeSpecificMarkup()
    {
        const string html = """
            <?xml version="1.0"?>
            <![if gte mso 9]>
            <o:p>ignore</o:p>
            <![endif]>
            <html><body><w:sdt><p>kept</p></w:sdt></body></html>
            """;

        var normalized = OfficeHtmlNormalizer.Normalize(html);

        Assert.DoesNotContain("<?xml", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<![if", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<o:p", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<w:sdt", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p>kept</p>", normalized, StringComparison.OrdinalIgnoreCase);
    }
}
