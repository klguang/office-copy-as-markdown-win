namespace OfficeCopyAsMarkdown;

internal sealed class AppSettings
{
    public HotkeySettings Hotkey { get; set; } = HotkeySettings.CreateDefault();

    public HotkeyGesture ResolveHotkey()
    {
        return Hotkey.TryToGesture(out var hotkey, out _)
            ? hotkey
            : HotkeyGesture.Default;
    }

    public static AppSettings CreateDefault()
    {
        return new AppSettings
        {
            Hotkey = HotkeySettings.CreateDefault()
        };
    }

    public static AppSettings FromHotkey(HotkeyGesture hotkey)
    {
        return new AppSettings
        {
            Hotkey = HotkeySettings.FromGesture(hotkey)
        };
    }
}

internal sealed class HotkeySettings
{
    public string[] Modifiers { get; set; } = ["Ctrl", "Shift"];

    public string Key { get; set; } = Keys.C.ToString();

    public bool TryToGesture(out HotkeyGesture hotkey, out string? error)
    {
        hotkey = HotkeyGesture.Default;

        var modifiers = Keys.None;
        foreach (var modifierName in Modifiers)
        {
            if (!HotkeyGesture.TryParseModifier(modifierName, out var modifier))
            {
                error = $"Unsupported modifier '{modifierName}'.";
                return false;
            }

            modifiers |= modifier;
        }

        if (!Enum.TryParse<Keys>(Key, ignoreCase: true, out var key))
        {
            error = $"Unsupported key '{Key}'.";
            return false;
        }

        return HotkeyGesture.TryCreate(modifiers, key, out hotkey, out error);
    }

    public static HotkeySettings CreateDefault() => FromGesture(HotkeyGesture.Default);

    public static HotkeySettings FromGesture(HotkeyGesture hotkey)
    {
        var modifiers = new List<string>(3);
        if (hotkey.HasModifier(Keys.Control))
        {
            modifiers.Add("Ctrl");
        }

        if (hotkey.HasModifier(Keys.Alt))
        {
            modifiers.Add("Alt");
        }

        if (hotkey.HasModifier(Keys.Shift))
        {
            modifiers.Add("Shift");
        }

        return new HotkeySettings
        {
            Modifiers = [.. modifiers],
            Key = hotkey.Key.ToString()
        };
    }
}
