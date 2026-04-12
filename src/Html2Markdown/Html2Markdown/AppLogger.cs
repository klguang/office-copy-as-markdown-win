namespace Html2Markdown;

internal enum LogLevel
{
    None,
    Error,
    Warning,
    Information,
    Debug,
    Trace
}

internal static class AppLogger
{
    public static bool IsEnabled => HtmlToMarkdownLog.IsEnabled;

    public static LogLevel CurrentLevel =>
        HtmlToMarkdownLog.IsEnabled
            ? LogLevel.Debug
            : LogLevel.None;

    public static void Error(string message) => HtmlToMarkdownLog.Error(message);

    public static void Warning(string message) => HtmlToMarkdownLog.Warning(message);

    public static void Info(string message) => HtmlToMarkdownLog.Information(message);

    public static void Debug(string message) => HtmlToMarkdownLog.Debug(message);
}
