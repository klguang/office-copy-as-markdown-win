using System.Diagnostics;

namespace OfficeCopyAsMarkdown.Services;

internal static class ForegroundOfficeDetector
{
    private static readonly HashSet<string> SupportedProcessNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "WINWORD",
        "ONENOTE"
    };

    public static bool TryGetSupportedForegroundProcess(out Process? process)
    {
        process = null;
        var handle = NativeMethods.GetForegroundWindow();
        if (handle == IntPtr.Zero)
        {
            AppLogger.Debug("No foreground window handle.");
            return false;
        }

        NativeMethods.GetWindowThreadProcessId(handle, out var processId);
        if (processId == 0)
        {
            AppLogger.Debug("Foreground window had no process id.");
            return false;
        }

        try
        {
            var current = Process.GetProcessById((int)processId);
            if (!SupportedProcessNames.Contains(current.ProcessName))
            {
                AppLogger.Debug($"Foreground process '{current.ProcessName}' is not supported.");
                current.Dispose();
                return false;
            }

            AppLogger.Debug($"Foreground process detected: {current.ProcessName} ({current.Id}).");
            process = current;
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error("Failed to inspect foreground process.", ex);
            return false;
        }
    }
}
