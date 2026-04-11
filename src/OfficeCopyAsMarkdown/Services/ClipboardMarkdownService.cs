using System.Drawing.Imaging;
using System.Text;

namespace OfficeCopyAsMarkdown.Services;

internal sealed class ClipboardMarkdownService
{
    private static readonly SemaphoreSlim Gate = new(1, 1);

    public async Task<MarkdownConversionResult> CopyForegroundSelectionAsMarkdownAsync()
    {
        if (!await Gate.WaitAsync(0))
        {
            AppLogger.Warning("Conversion request rejected because another conversion is in progress.");
            return MarkdownConversionResult.Fail("A conversion is already running.");
        }

        try
        {
            if (!ForegroundOfficeDetector.TryGetSupportedForegroundProcess(out var process))
            {
                AppLogger.Debug("Conversion aborted because the foreground app is not Word or OneNote.");
                return MarkdownConversionResult.Fail("Bring Word or OneNote to the foreground first.");
            }

            using (process)
            {
                var processName = process!.ProcessName;
                AppLogger.Debug($"Beginning clipboard conversion for {processName}.");
                var beforeSequence = NativeMethods.GetClipboardSequenceNumber();
                NativeMethods.SendCtrlC();
                AppLogger.Debug($"Sent Ctrl+C to {processName}. Clipboard sequence before copy: {beforeSequence}.");

                var snapshot = await WaitForStableClipboardSnapshotAsync(beforeSequence, TimeSpan.FromSeconds(2));
                if (snapshot is null)
                {
                    AppLogger.Warning($"Clipboard did not yield a stable usable snapshot after Ctrl+C in {processName}.");
                    return MarkdownConversionResult.Fail($"The selection in {processName} was not copied.");
                }

                AppLogger.Debug(
                    $"Clipboard snapshot captured. Html={snapshot.Html is not null}, TextLength={snapshot.Text?.Length ?? 0}, ImageBytes={snapshot.ImagePng?.Length ?? 0}, Signature={snapshot.StabilitySignature}.");
                var markdown = ConvertClipboardToMarkdown(snapshot);
                if (string.IsNullOrWhiteSpace(markdown))
                {
                    AppLogger.Warning("Clipboard snapshot could not be converted to Markdown.");
                    return MarkdownConversionResult.Fail("The selection could not be converted to Markdown.");
                }

                var dataObject = new DataObject();
                dataObject.SetText(markdown, TextDataFormat.UnicodeText);
                dataObject.SetText(markdown, TextDataFormat.Text);
                Clipboard.SetDataObject(dataObject, true);
                AppLogger.Debug($"Markdown written back to clipboard. Length={markdown.Length}.");

                return MarkdownConversionResult.Ok($"Converted {processName} selection to Markdown and preserved all detected text content.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("Clipboard conversion threw an exception.", ex);
            return MarkdownConversionResult.Fail($"Conversion failed: {ex.Message}");
        }
        finally
        {
            AppLogger.Debug("Conversion gate released.");
            Gate.Release();
        }
    }

    private static async Task<ClipboardSnapshot?> WaitForStableClipboardSnapshotAsync(uint initialSequence, TimeSpan timeout)
    {
        var started = DateTime.UtcNow;
        ClipboardSnapshot? latestUsableSnapshot = null;
        string? latestSignature = null;
        var stableReadCount = 0;

        while (DateTime.UtcNow - started < timeout)
        {
            if (NativeMethods.GetClipboardSequenceNumber() == initialSequence)
            {
                await Task.Delay(50);
                continue;
            }

            var snapshot = ReadClipboardSnapshot();
            if (snapshot.HasUsableData)
            {
                if (string.Equals(snapshot.StabilitySignature, latestSignature, StringComparison.Ordinal))
                {
                    stableReadCount++;
                    if (stableReadCount >= 1)
                    {
                        AppLogger.Debug($"Clipboard snapshot stabilized after copy. Signature={snapshot.StabilitySignature}.");
                        return snapshot;
                    }
                }
                else
                {
                    latestUsableSnapshot = snapshot;
                    latestSignature = snapshot.StabilitySignature;
                    stableReadCount = 0;
                }
            }

            await Task.Delay(50);
        }

        if (latestUsableSnapshot is not null)
        {
            AppLogger.Warning("Clipboard snapshot did not stabilize before timeout. Proceeding with the latest usable snapshot.");
            return latestUsableSnapshot;
        }

        AppLogger.Debug("Timed out waiting for clipboard data after Ctrl+C.");
        return null;
    }

    private static ClipboardSnapshot ReadClipboardSnapshot()
    {
        var snapshot = new ClipboardSnapshot();

        if (Clipboard.ContainsData(DataFormats.Html))
        {
            if (Clipboard.TryGetData<string>(DataFormats.Html, out var html))
            {
                snapshot.Html = html;
            }
        }

        if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
        {
            snapshot.Text = Clipboard.GetText(TextDataFormat.UnicodeText);
        }

        if (Clipboard.ContainsImage())
        {
            using var image = Clipboard.GetImage();
            if (image is not null)
            {
                using var stream = new MemoryStream();
                image.Save(stream, ImageFormat.Png);
                snapshot.ImagePng = stream.ToArray();
            }
        }

        return snapshot;
    }

    private static string ConvertClipboardToMarkdown(ClipboardSnapshot snapshot)
    {
        if (!string.IsNullOrWhiteSpace(snapshot.Html))
        {
            var fragment = CfHtmlExtractor.ExtractFragment(snapshot.Html!);
            AppLogger.Debug($"Extracted HTML fragment. Length={fragment.Length}.");
            HtmlStructureLogger.LogFragmentStructure(fragment);
            var markdown = MarkdownConverter.Convert(fragment, snapshot.ImagePng);
            if (!string.IsNullOrWhiteSpace(markdown))
            {
                var repaired = MarkdownContentGuard.RepairMarkdown(markdown, snapshot.Text);
                if (repaired.IsComplete)
                {
                    if (!string.Equals(repaired.Markdown, markdown, StringComparison.Ordinal))
                    {
                        AppLogger.Debug("Added missing source text back into the Markdown output.");
                    }

                    AppLogger.Debug("Converted clipboard HTML to complete Markdown.");
                    return repaired.Markdown;
                }

                AppLogger.Warning("Markdown output remained incomplete after repair. Falling back to conservative Markdown.");
            }

            if (!string.IsNullOrWhiteSpace(snapshot.Text))
            {
                var conservativeMarkdown = MarkdownContentGuard.BuildConservativeMarkdown(snapshot.Text!);
                var conservativeResult = MarkdownContentGuard.RepairMarkdown(conservativeMarkdown, snapshot.Text);
                if (conservativeResult.IsComplete)
                {
                    AppLogger.Debug("Using conservative Markdown derived from the complete source text.");
                    return conservativeResult.Markdown;
                }
            }
        }

        if (snapshot.ImagePng is { Length: > 0 })
        {
            AppLogger.Warning("Falling back to embedded PNG data URI because HTML conversion did not produce Markdown.");
            var dataUri = $"data:image/png;base64,{Convert.ToBase64String(snapshot.ImagePng)}";
            return $"![clipboard image]({dataUri})";
        }

        if (!string.IsNullOrWhiteSpace(snapshot.Text))
        {
            AppLogger.Debug("Falling back to conservative Markdown derived from plain text.");
            return MarkdownContentGuard.BuildConservativeMarkdown(snapshot.Text!);
        }

        AppLogger.Warning("Clipboard snapshot did not contain HTML, text, or image data that could be used.");
        return string.Empty;
    }

    private static string NormalizeText(string text)
    {
        var normalized = text.ReplaceLineEndings("\n").Trim();
        var lines = normalized.Split('\n')
            .Select(line => line.TrimEnd())
            .ToArray();

        var builder = new StringBuilder();
        for (var index = 0; index < lines.Length; index++)
        {
            builder.AppendLine(lines[index]);
        }

        return builder.ToString().Trim();
    }

    private sealed class ClipboardSnapshot
    {
        public string? Html { get; set; }

        public string? Text { get; set; }

        public byte[]? ImagePng { get; set; }

        public bool HasUsableData =>
            !string.IsNullOrWhiteSpace(Html) ||
            !string.IsNullOrWhiteSpace(Text) ||
            ImagePng is { Length: > 0 };

        public string StabilitySignature =>
            $"{Html?.Length ?? 0}:{Text?.Length ?? 0}:{ImagePng?.Length ?? 0}";
    }
}
