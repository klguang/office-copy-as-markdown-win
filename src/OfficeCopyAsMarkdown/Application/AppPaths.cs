namespace OfficeCopyAsMarkdown;

internal static class AppPaths
{
    public static readonly string DataDirectoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "OfficeCopyAsMarkdown");

    public static readonly string LogDirectoryPath = Path.Combine(DataDirectoryPath, "logs");

    public static readonly string SettingsFilePath = Path.Combine(DataDirectoryPath, "settings.json");
}
