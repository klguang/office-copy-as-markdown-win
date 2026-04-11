namespace OfficeCopyAsMarkdown;

internal sealed class HotkeyWindow : NativeWindow, IDisposable
{
    private readonly int _id;
    private bool _disposed;

    public event EventHandler? HotkeyPressed;

    public HotkeyWindow(uint modifiers, uint virtualKey)
    {
        _id = GetHashCode();
        CreateHandle(new CreateParams());

        if (!NativeMethods.RegisterHotKey(Handle, _id, modifiers, virtualKey))
        {
            AppLogger.Error("Failed to register hotkey Ctrl+Shift+C.");
            throw new InvalidOperationException("Unable to register Ctrl+Shift+C.");
        }

        AppLogger.Info("Registered hotkey Ctrl+Shift+C.");
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == NativeMethods.WM_HOTKEY && m.WParam.ToInt32() == _id)
        {
            AppLogger.Debug("WM_HOTKEY received.");
            HotkeyPressed?.Invoke(this, EventArgs.Empty);
        }

        base.WndProc(ref m);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        NativeMethods.UnregisterHotKey(Handle, _id);
        AppLogger.Info("Unregistered hotkey Ctrl+Shift+C.");
        DestroyHandle();
        _disposed = true;
    }
}
