using System.Text;

namespace OfficeCopyAsMarkdown;

internal static class AppLogger
{
    private const string EnvironmentVariableName = "OFFICE_COPY_AS_MARKDOWN_LOG_LEVEL";
    private static readonly Lock SyncRoot = new();
    private static readonly string LogDirectoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OfficeCopyAsMarkdown",
        "logs");

    public static LogLevel CurrentLevel { get; } = ResolveLogLevel();

    public static bool IsEnabled => CurrentLevel != LogLevel.None;

    public static string CurrentLogFilePath => Path.Combine(LogDirectoryPath, $"{DateTime.Now:yyyyMMdd}.log");

    public static void Info(string message) => Write(LogLevel.Information, message);

    public static void Debug(string message) => Write(LogLevel.Debug, message);

    public static void Warning(string message) => Write(LogLevel.Warning, message);

    public static void Error(string message, Exception? exception = null) => Write(LogLevel.Error, message, exception);

    public static void Write(LogLevel level, string message, Exception? exception = null)
    {
        if (!ShouldWrite(level))
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(LogDirectoryPath);

            var builder = new StringBuilder()
                .Append('[').Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")).Append("] ")
                .Append('[').Append(level).Append("] ")
                .Append(message);

            if (exception is not null)
            {
                builder.AppendLine()
                    .Append(exception);
            }

            var line = builder.AppendLine().ToString();
            lock (SyncRoot)
            {
                File.AppendAllText(CurrentLogFilePath, line, Encoding.UTF8);
            }
        }
        catch
        {
        }
    }

    private static bool ShouldWrite(LogLevel level) => level <= CurrentLevel && CurrentLevel != LogLevel.None;

    private static LogLevel ResolveLogLevel()
    {
        var configured = Environment.GetEnvironmentVariable(EnvironmentVariableName);
        if (Enum.TryParse<LogLevel>(configured, ignoreCase: true, out var parsedLevel))
        {
            return parsedLevel;
        }

#if DEBUG
        return LogLevel.Debug;
#else
        return LogLevel.None;
#endif
    }
}

internal enum LogLevel
{
    None = 0,
    Error = 1,
    Warning = 2,
    Information = 3,
    Debug = 4,
    Trace = 5
}
