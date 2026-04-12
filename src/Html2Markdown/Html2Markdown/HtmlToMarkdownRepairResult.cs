namespace Html2Markdown;

public sealed record HtmlToMarkdownRepairResult(bool IsComplete, string Markdown, IReadOnlyList<string> MissingLines);
