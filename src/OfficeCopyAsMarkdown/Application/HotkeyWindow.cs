namespace OfficeCopyAsMarkdown;

internal sealed class HotkeyWindow : NativeWindow, IDisposable
{
    private readonly int _id;
    private readonly IHotkeyRegistrar _registrar;
    private HotkeyGesture _activeHotkey;
    private bool _disposed;

    public event EventHandler? HotkeyPressed;

    public HotkeyWindow(HotkeyGesture hotkey)
        : this(hotkey, new NativeHotkeyRegistrar())
    {
    }

    internal HotkeyWindow(HotkeyGesture hotkey, IHotkeyRegistrar registrar)
    {
        _id = GetHashCode();
        _registrar = registrar;
        CreateHandle(new CreateParams());
        RegisterOrThrow(hotkey);
    }

    public HotkeyGesture ActiveHotkey => _activeHotkey;

    public bool TryUpdateHotkey(HotkeyGesture hotkey, out string? error)
    {
        if (hotkey == _activeHotkey)
        {
            error = null;
            return true;
        }

        var previousHotkey = _activeHotkey;
        _registrar.Unregister(Handle, _id);

        if (_registrar.Register(Handle, _id, hotkey))
        {
            _activeHotkey = hotkey;
            AppLogger.Info($"Registered hotkey {hotkey.DisplayText}.");
            error = null;
            return true;
        }

        AppLogger.Warning($"Failed to register hotkey {hotkey.DisplayText}. Restoring {previousHotkey.DisplayText}.");

        if (!_registrar.Register(Handle, _id, previousHotkey))
        {
            AppLogger.Error($"Failed to restore hotkey {previousHotkey.DisplayText} after a registration error.");
            throw new InvalidOperationException($"Unable to restore {previousHotkey.DisplayText} after registration failure.");
        }

        error = $"Unable to register {hotkey.DisplayText}. It may already be in use by another application.";
        return false;
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

        _registrar.Unregister(Handle, _id);
        AppLogger.Info($"Unregistered hotkey {_activeHotkey.DisplayText}.");
        DestroyHandle();
        _disposed = true;
    }

    private void RegisterOrThrow(HotkeyGesture hotkey)
    {
        if (!_registrar.Register(Handle, _id, hotkey))
        {
            AppLogger.Error($"Failed to register hotkey {hotkey.DisplayText}.");
            throw new InvalidOperationException($"Unable to register {hotkey.DisplayText}.");
        }

        _activeHotkey = hotkey;
        AppLogger.Info($"Registered hotkey {hotkey.DisplayText}.");
    }
}
