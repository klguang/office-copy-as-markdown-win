namespace OfficeCopyAsMarkdown;

internal interface IHotkeyRegistrar
{
    bool Register(IntPtr handle, int id, HotkeyGesture hotkey);

    bool Unregister(IntPtr handle, int id);
}

internal sealed class NativeHotkeyRegistrar : IHotkeyRegistrar
{
    public bool Register(IntPtr handle, int id, HotkeyGesture hotkey)
    {
        return NativeMethods.RegisterHotKey(handle, id, hotkey.NativeModifiers, hotkey.VirtualKey);
    }

    public bool Unregister(IntPtr handle, int id)
    {
        return NativeMethods.UnregisterHotKey(handle, id);
    }
}
