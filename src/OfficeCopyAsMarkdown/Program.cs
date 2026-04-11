namespace OfficeCopyAsMarkdown;

static class Program
{
    [STAThread]
    static void Main()
    {
        AppLogger.Info($"Application starting. LogLevel={AppLogger.CurrentLevel}");
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, args) =>
        {
            AppLogger.Error("Unhandled UI thread exception.", args.Exception);
            ShowFatalError(args.Exception);
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            var exception = args.ExceptionObject as Exception;
            AppLogger.Error("Unhandled AppDomain exception.", exception);
        };
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            AppLogger.Error("Unobserved task exception.", args.Exception);
            args.SetObserved();
        };

        ApplicationConfiguration.Initialize();
        try
        {
            Application.Run(new TrayApplicationContext());
        }
        catch (Exception ex)
        {
            AppLogger.Error("Fatal exception escaped the message loop.", ex);
            ShowFatalError(ex);
            throw;
        }
        finally
        {
            AppLogger.Info("Application shutting down.");
        }
    }

    private static void ShowFatalError(Exception exception)
    {
        var message = AppLogger.IsEnabled
            ? $"Office Copy as Markdown crashed. Log: {AppLogger.CurrentLogFilePath}"
            : "Office Copy as Markdown crashed.";

        MessageBox.Show(
            $"{message}{Environment.NewLine}{Environment.NewLine}{exception.Message}",
            "Office Copy as Markdown",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }
}
