namespace OfficeCopyAsMarkdown.Services;

internal sealed record MarkdownConversionResult(bool Success, string Message)
{
    public static MarkdownConversionResult Ok(string message) => new(true, message);

    public static MarkdownConversionResult Fail(string message) => new(false, message);
}
