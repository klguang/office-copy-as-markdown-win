using HtmlAgilityPack;

namespace Html2Markdown;

public sealed class DefaultHeadingInferenceStrategy : IHeadingInferenceStrategy
{
    private const int MaximumCandidateHeadingLength = 30;
    private static readonly char[] DisallowedHeadingTerminators = ['\u3002', '.', '\uFF0C', ',', '\uFF1B', ';'];
    private static readonly CandidateHeadingInferenceOptions CandidateHeadingInference = new(MaxLevels: 4, SparseStartLevel: 2);

    public static DefaultHeadingInferenceStrategy Instance { get; } = new();

    private DefaultHeadingInferenceStrategy()
    {
    }

    public IReadOnlyDictionary<HtmlNode, int> InferHeadingLevels(HtmlNode root, IHtmlDialectAdapter dialectAdapter)
    {
        var hasSemanticHeading = root.DescendantsAndSelf().Any(node => MarkdownConverter.TryGetSemanticHeadingLevel(node, dialectAdapter, out _));
        AppLogger.Debug($"Heading analysis: semantic heading present = {hasSemanticHeading}.");
        if (hasSemanticHeading)
        {
            AppLogger.Debug("Heading analysis: semantic heading detected, candidate heading inference disabled for the entire fragment.");
            return new Dictionary<HtmlNode, int>();
        }

        AppLogger.Debug($"Heading analysis: candidate heading max levels = {CandidateHeadingInference.EffectiveMaxLevels}, sparse start level = {CandidateHeadingInference.EffectiveSparseStartLevel}.");

        var bodyFontSamples = CollectBodyFontSizes(root);
        if (bodyFontSamples.Count == 0)
        {
            AppLogger.Debug("Heading analysis: no font-size samples found for the current fragment.");
            return new Dictionary<HtmlNode, int>();
        }

        var baseline = bodyFontSamples
            .GroupBy(size => size)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key)
            .First()
            .Key;
        AppLogger.Debug($"Heading analysis: body font baseline = {baseline:F2}pt.");

        var baseCandidates = CollectBaseCandidates(root, dialectAdapter);
        AppLogger.Debug($"Heading analysis: base candidate count = {baseCandidates.Count}.");
        if (baseCandidates.Count == 0)
        {
            return new Dictionary<HtmlNode, int>();
        }

        var effectiveCandidates = baseCandidates
            .Where(candidate =>
                candidate.FontSizePt >= baseline + 2d ||
                candidate.FontSizePt >= baseline * 1.15d ||
                (candidate.IsBold && candidate.FontSizePt >= baseline))
            .ToList();
        AppLogger.Debug($"Heading analysis: effective candidate count = {effectiveCandidates.Count}.");
        if (effectiveCandidates.Count == 0)
        {
            return new Dictionary<HtmlNode, int>();
        }

        var orderedFontBands = effectiveCandidates
            .Select(candidate => candidate.FontSizePt)
            .Distinct()
            .OrderByDescending(size => size)
            .Take(CandidateHeadingInference.EffectiveMaxLevels)
            .ToList();

        var levelOffset = orderedFontBands.Count < CandidateHeadingInference.EffectiveMaxLevels
            ? CandidateHeadingInference.EffectiveSparseStartLevel - 1
            : 0;

        var mappingDescription = orderedFontBands.Count == 0
            ? "none"
            : string.Join(", ", orderedFontBands.Select((size, levelIndex) => $"{size:F2}pt=>H{levelIndex + 1 + levelOffset}"));
        AppLogger.Debug($"Heading analysis: font bands = {mappingDescription}.");

        var nodeLevels = new Dictionary<HtmlNode, int>();
        foreach (var candidate in effectiveCandidates)
        {
            var bandIndex = orderedFontBands.FindIndex(size => Math.Abs(size - candidate.FontSizePt) < 0.01d);
            if (bandIndex >= 0)
            {
                nodeLevels[candidate.Node] = bandIndex + 1 + levelOffset;
            }
        }

        return nodeLevels;
    }

    private static List<double> CollectBodyFontSizes(HtmlNode root)
    {
        var sizes = new List<double>();

        foreach (var node in root.DescendantsAndSelf())
        {
            if (MarkdownConverter.ShouldSkip(node) || node.NodeType != HtmlNodeType.Element)
            {
                continue;
            }

            var text = MarkdownConverter.ExtractCandidateText(node);
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (MarkdownConverter.TryExtractFontSize(node, out var fontSizePt))
            {
                sizes.Add(Math.Round(fontSizePt, 2));
            }
        }

        return sizes;
    }

    private static List<HeadingCandidate> CollectBaseCandidates(HtmlNode root, IHtmlDialectAdapter dialectAdapter)
    {
        var candidates = new List<HeadingCandidate>();

        foreach (var node in root.Descendants())
        {
            if (!MarkdownConverter.IsCandidateBlock(node))
            {
                continue;
            }

            if (MarkdownConverter.HasBlockedAncestor(node, dialectAdapter))
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is inside a blocked container.");
                continue;
            }

            if (MarkdownConverter.IsStructuralContainer(node))
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a structural container.");
                continue;
            }

            if (node.SelectSingleNode(".//br") is not null)
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is not a single visual line.");
                continue;
            }

            if (dialectAdapter.ShouldBlockHeadingInference(node))
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is blocked by the active HTML dialect adapter.");
                continue;
            }

            if (MarkdownConverter.IsQuoteBlock(node, dialectAdapter) || MarkdownConverter.IsCodeBlock(node))
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a quote or code block.");
                continue;
            }

            var text = MarkdownConverter.ExtractCandidateText(node);
            if (string.IsNullOrWhiteSpace(text))
            {
                AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' text is empty.");
                continue;
            }

            if (text.Length > MaximumCandidateHeadingLength)
            {
                AppLogger.Debug($"Heading candidate rejected: '{text}' exceeds {MaximumCandidateHeadingLength} characters.");
                continue;
            }

            if (DisallowedHeadingTerminators.Contains(text[^1]))
            {
                AppLogger.Debug($"Heading candidate rejected: '{text}' ends with disallowed punctuation.");
                continue;
            }

            if (!MarkdownConverter.TryExtractFontSize(node, out var fontSizePt))
            {
                AppLogger.Debug($"Heading candidate rejected: '{text}' has no extractable font-size.");
                continue;
            }

            candidates.Add(new HeadingCandidate(node, Math.Round(fontSizePt, 2), MarkdownConverter.IsEntireTextBold(node)));
        }

        return candidates;
    }

    private sealed record CandidateHeadingInferenceOptions(int MaxLevels, int SparseStartLevel)
    {
        public int EffectiveMaxLevels => Math.Max(1, MaxLevels);
        public int EffectiveSparseStartLevel => Math.Clamp(SparseStartLevel, 1, 6);
    }

    private sealed record HeadingCandidate(HtmlNode Node, double FontSizePt, bool IsBold);
}
