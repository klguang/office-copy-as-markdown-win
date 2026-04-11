using System.Text.RegularExpressions;

namespace OfficeCopyAsMarkdown.Services;

internal static class OfficeHtmlNormalizer
{
    private static readonly Regex ConditionalRegex = new(
        @"<!\[(if|endif)[^\]]*\]>",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex OfficeNamespaceRegex = new(
        @"</?(o|w|v|st1):[^>]*>",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex XmlDeclarationsRegex = new(
        @"<\?xml[^>]*\?>",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static string Normalize(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        return XmlDeclarationsRegex.Replace(
            OfficeNamespaceRegex.Replace(
                ConditionalRegex.Replace(html, string.Empty),
                string.Empty),
            string.Empty);
    }
}
