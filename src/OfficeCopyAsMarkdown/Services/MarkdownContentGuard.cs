using System.Text.RegularExpressions;

namespace OfficeCopyAsMarkdown.Services;

internal sealed record MarkdownRepairResult(bool IsComplete, string Markdown, IReadOnlyList<string> MissingLines);

internal static class MarkdownContentGuard
{
    private static readonly Regex MarkdownImageRegex = new(@"!\[(?<alt>[^\]]*)\]\([^)]+\)", RegexOptions.Compiled);
    private static readonly Regex MarkdownLinkRegex = new(@"\[(?<text>[^\]]+)\]\([^)]+\)", RegexOptions.Compiled);
    private static readonly Regex MarkdownDecorationRegex = new(@"(\*\*|__|~~|`)", RegexOptions.Compiled);
    private static readonly Regex MarkdownPrefixRegex = new(@"^(>+\s*)?(#{1,6}\s+)?((\d+|[A-Za-z]+)[\.\)]\s+|[-*+]\s+)?(\[[ xX]\]\s+)?", RegexOptions.Compiled);
    private static readonly Regex SourcePrefixRegex = new(@"^((\d+|[A-Za-z]+)[\.\)]\s+|[\u2022\u00B7\u25CB\u25E6\u25CFo\-*]\s+|[\u2610\u2611\u2612]\s+)+", RegexOptions.Compiled);
    private static readonly Regex OrderedListRegex = new(@"^(?<marker>(\d+|[A-Za-z]+)[\.\)])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex BulletListRegex = new(@"^(?<marker>[\u2022\u00B7\u25CB\u25E6\u25CFo\-*])\s+(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex CheckboxRegex = new(@"^(?<box>[\u2610\u2611\u2612])\s*(?<text>.+)$", RegexOptions.Compiled);
    private static readonly Regex MarkdownHeadingRegex = new(@"^\s*#{1,6}\s+", RegexOptions.Compiled);
    private static readonly Regex MarkdownTableSeparatorRegex = new(@"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*:?-{3,}:?\s*\|?\s*$", RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    public static MarkdownRepairResult RepairMarkdown(string markdown, string? sourceText)
    {
        if (string.IsNullOrWhiteSpace(markdown) || string.IsNullOrWhiteSpace(sourceText))
        {
            return new MarkdownRepairResult(!string.IsNullOrWhiteSpace(markdown), markdown, []);
        }

        var sourceLines = ExtractSourceLines(sourceText);
        if (sourceLines.Count == 0)
        {
            return new MarkdownRepairResult(true, markdown, []);
        }

        var initial = AnalyzeCoverage(sourceLines, markdown);
        LogCoverage("initial", initial);
        if (initial.MissingSourceLineIndexes.Count == 0)
        {
            return new MarkdownRepairResult(true, markdown, []);
        }

        var repairedMarkdown = RecoverMissingSourceLines(markdown, sourceLines, initial);
        var repaired = AnalyzeCoverage(sourceLines, repairedMarkdown);
        LogCoverage("repaired", repaired);

        var missingLines = repaired.MissingSourceLineIndexes
            .Select(index => sourceLines[index].RawText)
            .ToArray();

        return new MarkdownRepairResult(repaired.MissingSourceLineIndexes.Count == 0, repairedMarkdown, missingLines);
    }

    public static string BuildConservativeMarkdown(string sourceText)
    {
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            return string.Empty;
        }

        var sourceLines = ExtractSourceLines(sourceText);
        return string.Join('\n', sourceLines.Select(RenderSourceLine)).Trim();
    }

    public static bool ShouldKeepMarkdown(string markdown, string? sourceText)
    {
        if (string.IsNullOrWhiteSpace(markdown) || string.IsNullOrWhiteSpace(sourceText))
        {
            return !string.IsNullOrWhiteSpace(markdown);
        }

        var sourceLines = ExtractSourceLines(sourceText);
        if (sourceLines.Count == 0)
        {
            return true;
        }

        var analysis = AnalyzeCoverage(sourceLines, markdown);
        LogCoverage("initial", analysis);
        return analysis.MissingSourceLineIndexes.Count == 0;
    }

    private static CoverageAnalysis AnalyzeCoverage(IReadOnlyList<SourceLine> sourceLines, string markdown)
    {
        var markdownLines = ExtractComparableMarkdownLines(markdown);
        var matchedMarkdownLineIndexesBySource = Enumerable.Repeat(-1, sourceLines.Count).ToArray();
        var markdownSearchStart = 0;

        for (var sourceIndex = 0; sourceIndex < sourceLines.Count; sourceIndex++)
        {
            var comparableText = sourceLines[sourceIndex].ComparableText;
            for (var markdownIndex = markdownSearchStart; markdownIndex < markdownLines.Count; markdownIndex++)
            {
                if (!string.Equals(markdownLines[markdownIndex].ComparableText, comparableText, StringComparison.Ordinal))
                {
                    continue;
                }

                matchedMarkdownLineIndexesBySource[sourceIndex] = markdownLines[markdownIndex].LineIndex;
                markdownSearchStart = markdownIndex + 1;
                break;
            }
        }

        var missingSourceLineIndexes = matchedMarkdownLineIndexesBySource
            .Select((markdownLineIndex, sourceIndex) => new { markdownLineIndex, sourceIndex })
            .Where(item => item.markdownLineIndex < 0)
            .Select(item => item.sourceIndex)
            .ToArray();

        return new CoverageAnalysis(sourceLines, markdownLines, matchedMarkdownLineIndexesBySource, missingSourceLineIndexes);
    }

    private static void LogCoverage(string stage, CoverageAnalysis analysis)
    {
        var sample = analysis.MissingSourceLineIndexes
            .Take(3)
            .Select(index => analysis.SourceLines[index].RawText)
            .ToArray();
        var sampleText = sample.Length == 0 ? "none" : string.Join(" | ", sample);

        AppLogger.Debug(
            $"Markdown guard ({stage}): sourceLines={analysis.SourceLines.Count}, markdownLines={analysis.MarkdownLines.Count}, matched={analysis.MatchedLineCount}, coverage={analysis.CoverageRatio:F2}, lengthRatio={analysis.LengthRatio:F2}, missingSample={sampleText}.");
    }

    private static string RecoverMissingSourceLines(string markdown, IReadOnlyList<SourceLine> sourceLines, CoverageAnalysis analysis)
    {
        var markdownLines = markdown.ReplaceLineEndings("\n").Split('\n').ToList();
        if (markdownLines.Count == 1 && markdownLines[0].Length == 0)
        {
            markdownLines.Clear();
        }

        var sourcePositions = analysis.MatchedMarkdownLineIndexesBySource.ToArray();

        foreach (var missingSourceIndex in analysis.MissingSourceLineIndexes)
        {
            var insertionIndex = ResolveInsertionIndex(missingSourceIndex, sourcePositions, markdownLines.Count);
            var insertion = InsertSourceLine(markdownLines, insertionIndex, sourceLines[missingSourceIndex]);

            for (var index = 0; index < sourcePositions.Length; index++)
            {
                if (sourcePositions[index] >= insertion.InsertionIndex)
                {
                    sourcePositions[index] += insertion.InsertedLineCount;
                }
            }

            sourcePositions[missingSourceIndex] = insertion.ContentLineIndex;
        }

        return string.Join("\n", CollapseBlankLines(markdownLines)).Trim();
    }

    private static InsertionResult InsertSourceLine(List<string> markdownLines, int insertionIndex, SourceLine sourceLine)
    {
        var rendered = RenderSourceLine(sourceLine);
        var targetIndex = Math.Clamp(insertionIndex, 0, markdownLines.Count);
        var insertedLineCount = 0;

        if (sourceLine.Kind == SourceLineKind.Paragraph)
        {
            if (targetIndex > 0 && !string.IsNullOrWhiteSpace(markdownLines[targetIndex - 1]))
            {
                markdownLines.Insert(targetIndex, string.Empty);
                targetIndex++;
                insertedLineCount++;
            }

            markdownLines.Insert(targetIndex, rendered);
            insertedLineCount++;
            var contentLineIndex = targetIndex;
            targetIndex++;

            if (targetIndex < markdownLines.Count && !string.IsNullOrWhiteSpace(markdownLines[targetIndex]))
            {
                markdownLines.Insert(targetIndex, string.Empty);
                insertedLineCount++;
            }

            return new InsertionResult(insertionIndex, insertedLineCount, contentLineIndex);
        }

        if (targetIndex > 0 &&
            !string.IsNullOrWhiteSpace(markdownLines[targetIndex - 1]) &&
            !MarkdownHeadingRegex.IsMatch(markdownLines[targetIndex - 1]) &&
            !OrderedListRegex.IsMatch(markdownLines[targetIndex - 1].TrimStart()) &&
            !BulletListRegex.IsMatch(markdownLines[targetIndex - 1].TrimStart()) &&
            !markdownLines[targetIndex - 1].TrimStart().StartsWith("- [", StringComparison.Ordinal))
        {
            markdownLines.Insert(targetIndex, string.Empty);
            targetIndex++;
            insertedLineCount++;
        }

        markdownLines.Insert(targetIndex, rendered);
        insertedLineCount++;
        return new InsertionResult(insertionIndex, insertedLineCount, targetIndex);
    }

    private static int ResolveInsertionIndex(int sourceIndex, IReadOnlyList<int> sourcePositions, int markdownLineCount)
    {
        for (var previousIndex = sourceIndex - 1; previousIndex >= 0; previousIndex--)
        {
            if (sourcePositions[previousIndex] >= 0)
            {
                return sourcePositions[previousIndex] + 1;
            }
        }

        for (var nextIndex = sourceIndex + 1; nextIndex < sourcePositions.Count; nextIndex++)
        {
            if (sourcePositions[nextIndex] >= 0)
            {
                return sourcePositions[nextIndex];
            }
        }

        return markdownLineCount;
    }

    private static IReadOnlyList<string> CollapseBlankLines(IEnumerable<string> lines)
    {
        var collapsed = new List<string>();
        var previousWasBlank = false;

        foreach (var line in lines)
        {
            var isBlank = string.IsNullOrWhiteSpace(line);
            if (isBlank && previousWasBlank)
            {
                continue;
            }

            collapsed.Add(isBlank ? string.Empty : line.TrimEnd());
            previousWasBlank = isBlank;
        }

        return collapsed;
    }

    private static List<ComparableLine> ExtractComparableMarkdownLines(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            return [];
        }

        var normalized = markdown.ReplaceLineEndings("\n");
        normalized = MarkdownImageRegex.Replace(normalized, "${alt}");
        normalized = MarkdownLinkRegex.Replace(normalized, "${text}");
        normalized = MarkdownDecorationRegex.Replace(normalized, string.Empty);

        var lines = new List<ComparableLine>();
        var rawLines = normalized.Split('\n');
        for (var index = 0; index < rawLines.Length; index++)
        {
            var comparableText = NormalizeComparableLine(rawLines[index], isMarkdown: true);
            if (!string.IsNullOrWhiteSpace(comparableText))
            {
                lines.Add(new ComparableLine(index, comparableText));
            }
        }

        return lines;
    }

    private static List<SourceLine> ExtractSourceLines(string sourceText)
    {
        var lines = new List<SourceLine>();
        foreach (var rawLine in sourceText.ReplaceLineEndings("\n").Split('\n'))
        {
            var trimmed = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                continue;
            }

            if (CheckboxRegex.Match(trimmed) is { Success: true } checkboxMatch)
            {
                var content = checkboxMatch.Groups["text"].Value.Trim();
                var comparable = NormalizeComparableText(content);
                var isChecked = checkboxMatch.Groups["box"].Value is "\u2611" or "\u2612";
                lines.Add(new SourceLine(trimmed, comparable, SourceLineKind.Task, string.Empty, content, isChecked));
                continue;
            }

            if (OrderedListRegex.Match(trimmed) is { Success: true } orderedMatch)
            {
                var marker = orderedMatch.Groups["marker"].Value;
                var content = orderedMatch.Groups["text"].Value.Trim();
                lines.Add(new SourceLine(trimmed, NormalizeComparableText(content), SourceLineKind.Ordered, marker, content, IsChecked: false));
                continue;
            }

            if (BulletListRegex.Match(trimmed) is { Success: true } bulletMatch)
            {
                var content = bulletMatch.Groups["text"].Value.Trim();
                lines.Add(new SourceLine(trimmed, NormalizeComparableText(content), SourceLineKind.Bullet, "-", content, IsChecked: false));
                continue;
            }

            lines.Add(new SourceLine(trimmed, NormalizeComparableText(trimmed), SourceLineKind.Paragraph, string.Empty, trimmed, IsChecked: false));
        }

        return lines;
    }

    private static string RenderSourceLine(SourceLine sourceLine)
    {
        return sourceLine.Kind switch
        {
            SourceLineKind.Task => $"- [{(sourceLine.IsChecked ? "x" : " ")}] {sourceLine.Content}",
            SourceLineKind.Ordered => $"{sourceLine.Marker} {sourceLine.Content}",
            SourceLineKind.Bullet => $"- {sourceLine.Content}",
            _ => sourceLine.Content
        };
    }

    private static string NormalizeComparableLine(string text, bool isMarkdown)
    {
        var trimmed = text.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        if (isMarkdown && LooksLikeMarkdownTableRow(trimmed))
        {
            return NormalizeMarkdownTableRow(trimmed);
        }

        trimmed = isMarkdown
            ? MarkdownPrefixRegex.Replace(trimmed, string.Empty)
            : SourcePrefixRegex.Replace(trimmed, string.Empty);

        return NormalizeComparableText(trimmed);
    }

    private static string NormalizeComparableText(string text)
    {
        var normalized = text.Trim()
            .Replace('\t', ' ')
            .Replace("→", "->", StringComparison.Ordinal)
            .Replace("—", "-", StringComparison.Ordinal)
            .Replace("＋", "+", StringComparison.Ordinal);

        normalized = WhitespaceRegex.Replace(normalized, " ");
        normalized = Regex.Replace(normalized, @"\s*(->)\s*", " $1 ");
        normalized = Regex.Replace(normalized, @"\s*([/:：+])\s*", "$1");
        normalized = Regex.Replace(normalized, @"(?<=[\u4E00-\u9FFF])\s+(?=[A-Za-z0-9])", string.Empty);
        normalized = Regex.Replace(normalized, @"(?<=[A-Za-z0-9])\s+(?=[\u4E00-\u9FFF])", string.Empty);
        normalized = Regex.Replace(normalized, @"(?<=[\u4E00-\u9FFF])\s+(?=[\u4E00-\u9FFF])", string.Empty);
        normalized = WhitespaceRegex.Replace(normalized, " ");
        return normalized.Trim();
    }

    private static bool LooksLikeMarkdownTableRow(string line) => line.Contains('|');

    private static string NormalizeMarkdownTableRow(string line)
    {
        if (MarkdownTableSeparatorRegex.IsMatch(line))
        {
            return string.Empty;
        }

        var content = line.Trim();
        if (content.StartsWith('|'))
        {
            content = content[1..];
        }

        if (content.EndsWith('|'))
        {
            content = content[..^1];
        }

        var cells = SplitMarkdownTableCells(content)
            .Select(cell => cell
                .Replace(@"\|", "|", StringComparison.Ordinal)
                .Replace("<br>", " ", StringComparison.OrdinalIgnoreCase)
                .Trim())
            .Where(cell => !string.IsNullOrWhiteSpace(cell));

        return NormalizeComparableText(string.Join(" ", cells));
    }

    private static IReadOnlyList<string> SplitMarkdownTableCells(string row)
    {
        var cells = new List<string>();
        var builder = new System.Text.StringBuilder();
        var isEscaped = false;

        foreach (var character in row)
        {
            if (isEscaped)
            {
                builder.Append(character);
                isEscaped = false;
                continue;
            }

            if (character == '\\')
            {
                builder.Append(character);
                isEscaped = true;
                continue;
            }

            if (character == '|')
            {
                cells.Add(builder.ToString());
                builder.Clear();
                continue;
            }

            builder.Append(character);
        }

        cells.Add(builder.ToString());
        return cells;
    }

    private sealed record SourceLine(string RawText, string ComparableText, SourceLineKind Kind, string Marker, string Content, bool IsChecked);

    private sealed record ComparableLine(int LineIndex, string ComparableText);

    private sealed record CoverageAnalysis(
        IReadOnlyList<SourceLine> SourceLines,
        IReadOnlyList<ComparableLine> MarkdownLines,
        IReadOnlyList<int> MatchedMarkdownLineIndexesBySource,
        IReadOnlyList<int> MissingSourceLineIndexes)
    {
        public int MatchedLineCount => SourceLines.Count - MissingSourceLineIndexes.Count;

        public double CoverageRatio => SourceLines.Count == 0 ? 1d : MatchedLineCount / (double)SourceLines.Count;

        public double LengthRatio
        {
            get
            {
                var sourceLength = string.Join("\n", SourceLines.Select(line => line.ComparableText)).Length;
                var markdownLength = string.Join("\n", MarkdownLines.Select(line => line.ComparableText)).Length;
                return markdownLength / (double)Math.Max(1, sourceLength);
            }
        }
    }

    private sealed record InsertionResult(int InsertionIndex, int InsertedLineCount, int ContentLineIndex);

    private enum SourceLineKind
    {
        Paragraph,
        Bullet,
        Ordered,
        Task
    }
}
