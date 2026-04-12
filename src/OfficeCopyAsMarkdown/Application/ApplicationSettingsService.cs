using System.Text.Json;

namespace OfficeCopyAsMarkdown;

internal sealed class ApplicationSettingsService
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    public ApplicationSettingsService(string? settingsFilePath = null)
    {
        SettingsFilePath = settingsFilePath ?? AppPaths.SettingsFilePath;
    }

    public string SettingsFilePath { get; }

    public AppSettings Load()
    {
        if (!File.Exists(SettingsFilePath))
        {
            AppLogger.Info($"Settings file not found. Using defaults at {SettingsFilePath}.");
            return AppSettings.CreateDefault();
        }

        try
        {
            var json = File.ReadAllText(SettingsFilePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions);
            string? error = null;
            if (settings?.Hotkey.TryToGesture(out _, out error) == true)
            {
                AppLogger.Info($"Loaded settings from {SettingsFilePath}.");
                return settings;
            }

            AppLogger.Warning($"Settings file is invalid. Using defaults. {error}");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            AppLogger.Warning($"Failed to load settings from {SettingsFilePath}. Using defaults. {ex.Message}");
        }

        return AppSettings.CreateDefault();
    }

    public void Save(AppSettings settings)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath)!);

        var normalized = settings.Hotkey.TryToGesture(out var hotkey, out var error)
            ? AppSettings.FromHotkey(hotkey)
            : throw new InvalidOperationException(error ?? "Settings are invalid.");

        var json = JsonSerializer.Serialize(normalized, SerializerOptions);
        File.WriteAllText(SettingsFilePath, json);
        AppLogger.Info($"Saved settings to {SettingsFilePath}.");
    }
}
