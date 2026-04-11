using OfficeCopyAsMarkdown.Services;
using System.Diagnostics;

namespace OfficeCopyAsMarkdown;

internal sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _notifyIcon;
    private readonly HotkeyWindow _hotkeyWindow;
    private readonly ClipboardMarkdownService _clipboardMarkdownService;

    public TrayApplicationContext()
    {
        AppLogger.Info("Initializing tray application context.");
        _clipboardMarkdownService = new ClipboardMarkdownService();
        _hotkeyWindow = new HotkeyWindow(NativeMethods.MOD_CONTROL | NativeMethods.MOD_SHIFT, (uint)Keys.C);
        _hotkeyWindow.HotkeyPressed += OnHotkeyPressed;

        var contextMenu = new ContextMenuStrip();
        contextMenu.Items.Add("Copy Selection as Markdown", null, async (_, _) => await ConvertCurrentSelectionAsync());
        if (AppLogger.IsEnabled)
        {
            contextMenu.Items.Add("Open Log Folder", null, (_, _) => OpenLogFolder());
        }
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Exit", null, (_, _) => ExitThread());

        _notifyIcon = new NotifyIcon
        {
            Text = "Office Copy as Markdown",
            Icon = SystemIcons.Application,
            Visible = true,
            ContextMenuStrip = contextMenu
        };

        _notifyIcon.DoubleClick += async (_, _) => await ConvertCurrentSelectionAsync();
        _notifyIcon.ShowBalloonTip(
            2500,
            "Office Copy as Markdown",
            "Active for Word and OneNote when their window is in the foreground. Shortcut: Ctrl+Shift+C",
            ToolTipIcon.Info);
        AppLogger.Info("Tray icon ready.");
    }

    protected override void ExitThreadCore()
    {
        AppLogger.Info("Exit requested.");
        _hotkeyWindow.HotkeyPressed -= OnHotkeyPressed;
        _hotkeyWindow.Dispose();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        base.ExitThreadCore();
    }

    private async void OnHotkeyPressed(object? sender, EventArgs e)
    {
        AppLogger.Debug("Hotkey pressed.");
        await ConvertCurrentSelectionAsync();
    }

    private async Task ConvertCurrentSelectionAsync()
    {
        AppLogger.Debug("Starting conversion from tray command.");
        var result = await _clipboardMarkdownService.CopyForegroundSelectionAsMarkdownAsync();

        if (!result.Success)
        {
            AppLogger.Warning($"Conversion failed: {result.Message}");
            _notifyIcon.ShowBalloonTip(2500, "Office Copy as Markdown", result.Message, ToolTipIcon.Warning);
            return;
        }

        AppLogger.Info(result.Message);
        _notifyIcon.ShowBalloonTip(2000, "Markdown copied", result.Message, ToolTipIcon.Info);
    }

    private static void OpenLogFolder()
    {
        var directory = Path.GetDirectoryName(AppLogger.CurrentLogFilePath);
        if (string.IsNullOrWhiteSpace(directory))
        {
            return;
        }

        Directory.CreateDirectory(directory);
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = directory,
            UseShellExecute = true
        });
    }
}
