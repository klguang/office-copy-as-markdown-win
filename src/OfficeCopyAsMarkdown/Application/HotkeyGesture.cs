using System.ComponentModel;

namespace OfficeCopyAsMarkdown;

internal readonly record struct HotkeyGesture(Keys Modifiers, Keys Key)
{
    private const Keys SupportedModifiers = Keys.Control | Keys.Alt | Keys.Shift;

    public static HotkeyGesture Default => new(Keys.Control | Keys.Shift, Keys.C);

    public Keys KeyData => (Modifiers & SupportedModifiers) | (Key & Keys.KeyCode);

    public string DisplayText => Format(KeyData);

    public uint NativeModifiers
    {
        get
        {
            var modifiers = 0U;
            if (HasModifier(Keys.Control))
            {
                modifiers |= NativeMethods.MOD_CONTROL;
            }

            if (HasModifier(Keys.Shift))
            {
                modifiers |= NativeMethods.MOD_SHIFT;
            }

            if (HasModifier(Keys.Alt))
            {
                modifiers |= NativeMethods.MOD_ALT;
            }

            return modifiers;
        }
    }

    public uint VirtualKey => (uint)(Key & Keys.KeyCode);

    public bool HasModifier(Keys modifier) => (Modifiers & modifier) == modifier;

    public static bool TryCreate(Keys keyData, out HotkeyGesture hotkey, out string? error)
    {
        return TryCreate(keyData & SupportedModifiers, keyData & Keys.KeyCode, out hotkey, out error);
    }

    public static bool TryCreate(Keys modifiers, Keys key, out HotkeyGesture hotkey, out string? error)
    {
        hotkey = Default;

        modifiers &= SupportedModifiers;
        key &= Keys.KeyCode;

        if (modifiers == Keys.None)
        {
            error = "Shortcut must include Ctrl, Alt, or Shift.";
            return false;
        }

        if (key == Keys.None || IsModifierKey(key))
        {
            error = "Shortcut must include a non-modifier key.";
            return false;
        }

        hotkey = new HotkeyGesture(modifiers, key);
        error = null;
        return true;
    }

    public static bool TryParseModifier(string value, out Keys modifier)
    {
        modifier = value.Trim().ToUpperInvariant() switch
        {
            "CTRL" or "CONTROL" => Keys.Control,
            "ALT" => Keys.Alt,
            "SHIFT" => Keys.Shift,
            _ => Keys.None
        };

        return modifier != Keys.None;
    }

    public static bool IsModifierKey(Keys key)
    {
        return (key & Keys.KeyCode) is Keys.ShiftKey or Keys.ControlKey or Keys.Menu;
    }

    private static string Format(Keys keyData)
    {
        var converter = TypeDescriptor.GetConverter(typeof(Keys));
        return converter.ConvertToInvariantString(keyData) ?? keyData.ToString();
    }
}
