using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OfficeCopyAsMarkdown;

internal static partial class NativeMethods
{
    public const int WM_HOTKEY = 0x0312;
    public const uint MOD_ALT = 0x0001;
    public const uint MOD_CONTROL = 0x0002;
    public const uint MOD_SHIFT = 0x0004;
    public const uint KEYEVENTF_KEYUP = 0x0002;
    public const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool UnregisterHotKey(IntPtr hWnd, int id);

    [LibraryImport("user32.dll")]
    public static partial IntPtr GetForegroundWindow();

    [LibraryImport("user32.dll")]
    public static partial uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [LibraryImport("user32.dll")]
    public static partial uint GetClipboardSequenceNumber();

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial void keybd_event(byte bVk, byte bScan, uint dwFlags, nuint dwExtraInfo);

    public static void SendCtrlC()
    {
        const byte control = (byte)Keys.ControlKey;
        const byte shift = (byte)Keys.ShiftKey;
        const byte c = (byte)Keys.C;

        AppLogger.Debug("Sending synthetic Ctrl+C with keybd_event.");

        keybd_event(shift, 0, KEYEVENTF_KEYUP, 0);
        Thread.Sleep(10);
        keybd_event(control, 0, 0, 0);
        keybd_event(c, 0, 0, 0);
        keybd_event(c, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(control, 0, KEYEVENTF_KEYUP, 0);
    }
}
