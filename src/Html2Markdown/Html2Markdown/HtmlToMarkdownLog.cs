using System.Threading;

namespace Html2Markdown;

internal static class HtmlToMarkdownLog
{
    private static readonly AsyncLocal<Action<HtmlToMarkdownLogLevel, string>?> CurrentSink = new();

    public static bool IsEnabled => CurrentSink.Value is not null;

    public static bool IsEnabledAt(HtmlToMarkdownLogLevel _) => IsEnabled;

    public static T RunWith<T>(Action<HtmlToMarkdownLogLevel, string>? sink, Func<T> action)
    {
        var previous = CurrentSink.Value;
        CurrentSink.Value = sink;

        try
        {
            return action();
        }
        finally
        {
            CurrentSink.Value = previous;
        }
    }

    public static void Error(string message) => Write(HtmlToMarkdownLogLevel.Error, message);

    public static void Warning(string message) => Write(HtmlToMarkdownLogLevel.Warning, message);

    public static void Information(string message) => Write(HtmlToMarkdownLogLevel.Information, message);

    public static void Debug(string message) => Write(HtmlToMarkdownLogLevel.Debug, message);

    public static void Trace(string message) => Write(HtmlToMarkdownLogLevel.Trace, message);

    private static void Write(HtmlToMarkdownLogLevel level, string message)
    {
        CurrentSink.Value?.Invoke(level, message);
    }
}
