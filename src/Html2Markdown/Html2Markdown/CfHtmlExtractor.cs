using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Html2Markdown;

internal static class CfHtmlExtractor
{
    private static readonly Regex HeaderRegex = new(
        @"StartFragment:(?<start>\d+).*?EndFragment:(?<end>\d+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

    public static string ExtractFragment(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var normalized = html.Replace("\0", string.Empty);
        var match = HeaderRegex.Match(normalized);
        if (!match.Success)
        {
            return StripMarkers(normalized);
        }

        var start = int.Parse(match.Groups["start"].Value, CultureInfo.InvariantCulture);
        var end = int.Parse(match.Groups["end"].Value, CultureInfo.InvariantCulture);

        if (start < 0 || end <= start || end > normalized.Length)
        {
            var fallback = ExtractByUtf8Offsets(normalized, start, end);
            return string.IsNullOrWhiteSpace(fallback) ? StripMarkers(normalized) : StripMarkers(fallback);
        }

        var byChars = normalized[start..end];
        var byBytes = ExtractByUtf8Offsets(normalized, start, end);
        return StripMarkers(string.IsNullOrWhiteSpace(byBytes) ? byChars : byBytes);
    }

    private static string StripMarkers(string html)
    {
        return html
            .Replace("<!--StartFragment-->", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("<!--EndFragment-->", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private static string ExtractByUtf8Offsets(string html, int start, int end)
    {
        var bytes = Encoding.UTF8.GetBytes(html);
        if (start < 0 || end <= start || end > bytes.Length)
        {
            return string.Empty;
        }

        return Encoding.UTF8.GetString(bytes[start..end]);
    }
}
