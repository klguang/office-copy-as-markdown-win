using OfficeCopyAsMarkdown.Services;
using System.Diagnostics;

namespace OfficeCopyAsMarkdown;

internal sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _notifyIcon;
    private readonly HotkeyWindow _hotkeyWindow;
    private readonly ClipboardMarkdownService _clipboardMarkdownService;
    private readonly ApplicationSettingsService _settingsService;
    private readonly ToolStripMenuItem _copyMenuItem;
    private AppSettings _settings;

    public TrayApplicationContext()
    {
        AppLogger.Info("Initializing tray application context.");
        _clipboardMarkdownService = new ClipboardMarkdownService();
        _settingsService = new ApplicationSettingsService();
        _settings = _settingsService.Load();

        var hotkey = _settings.ResolveHotkey();
        _hotkeyWindow = new HotkeyWindow(hotkey);
        _hotkeyWindow.HotkeyPressed += OnHotkeyPressed;

        var contextMenu = new ContextMenuStrip();
        _copyMenuItem = new ToolStripMenuItem(GetCopyCommandText(hotkey), null, async (_, _) => await ConvertCurrentSelectionAsync());
        contextMenu.Items.Add(_copyMenuItem);
        contextMenu.Items.Add("Settings...", null, (_, _) => OpenSettings());
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
            $"Active for Word and OneNote when their window is in the foreground. Shortcut: {hotkey.DisplayText}",
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

    private void OpenSettings()
    {
        using var form = new SettingsForm(_hotkeyWindow.ActiveHotkey);
        if (form.ShowDialog() != DialogResult.OK)
        {
            return;
        }

        var previousHotkey = _hotkeyWindow.ActiveHotkey;
        var newHotkey = form.SelectedHotkey;
        if (newHotkey == previousHotkey)
        {
            return;
        }

        if (!_hotkeyWindow.TryUpdateHotkey(newHotkey, out var error))
        {
            MessageBox.Show(
                error,
                "Office Copy as Markdown",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var updatedSettings = AppSettings.FromHotkey(newHotkey);
            _settingsService.Save(updatedSettings);
            _settings = updatedSettings;
            _copyMenuItem.Text = GetCopyCommandText(newHotkey);
            _notifyIcon.ShowBalloonTip(
                2000,
                "Shortcut updated",
                $"The new shortcut is {newHotkey.DisplayText}.",
                ToolTipIcon.Info);
        }
        catch (Exception ex)
        {
            AppLogger.Error("Failed to save updated hotkey settings.", ex);
            _hotkeyWindow.TryUpdateHotkey(previousHotkey, out _);
            _settings = AppSettings.FromHotkey(previousHotkey);
            _copyMenuItem.Text = GetCopyCommandText(previousHotkey);

            MessageBox.Show(
                "The shortcut was not saved. The previous shortcut is still active.",
                "Office Copy as Markdown",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
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

    private static string GetCopyCommandText(HotkeyGesture hotkey)
    {
        return $"Copy Selection as Markdown ({hotkey.DisplayText})";
    }
}
