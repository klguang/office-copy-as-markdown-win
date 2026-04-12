using HtmlAgilityPack;

namespace OfficeCopyAsMarkdown.Services;

internal static partial class MarkdownConverter
{
    private sealed class HeadingInference
    {
        private readonly Dictionary<HtmlNode, int> _nodeLevels;

        private HeadingInference(Dictionary<HtmlNode, int> nodeLevels)
        {
            _nodeLevels = nodeLevels;
        }

        public static HeadingInference Analyze(HtmlNode root, CandidateHeadingInferenceOptions candidateHeadingInference)
        {
            var hasSemanticHeading = root.DescendantsAndSelf().Any(node => TryGetSemanticHeadingLevel(node, out _));
            AppLogger.Debug($"Heading analysis: semantic heading present = {hasSemanticHeading}.");
            if (hasSemanticHeading)
            {
                AppLogger.Debug("Heading analysis: semantic heading detected, candidate heading inference disabled for the entire fragment.");
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            AppLogger.Debug($"Heading analysis: candidate heading max levels = {candidateHeadingInference.EffectiveMaxLevels}, sparse start level = {candidateHeadingInference.EffectiveSparseStartLevel}.");

            var bodyFontSamples = CollectBodyFontSizes(root);
            if (bodyFontSamples.Count == 0)
            {
                AppLogger.Debug("Heading analysis: no font-size samples found for the current fragment.");
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var baseline = bodyFontSamples
                .GroupBy(size => size)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key)
                .First()
                .Key;
            AppLogger.Debug($"Heading analysis: body font baseline = {baseline:F2}pt.");

            var baseCandidates = CollectBaseCandidates(root);
            AppLogger.Debug($"Heading analysis: base candidate count = {baseCandidates.Count}.");
            if (baseCandidates.Count == 0)
            {
                return new HeadingInference(new Dictionary<HtmlNode, int>());
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
                return new HeadingInference(new Dictionary<HtmlNode, int>());
            }

            var orderedFontBands = effectiveCandidates
                .Select(candidate => candidate.FontSizePt)
                .Distinct()
                .OrderByDescending(size => size)
                .Take(candidateHeadingInference.EffectiveMaxLevels)
                .ToList();

            var levelOffset = orderedFontBands.Count < candidateHeadingInference.EffectiveMaxLevels
                ? candidateHeadingInference.EffectiveSparseStartLevel - 1
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

            return new HeadingInference(nodeLevels);
        }

        public bool TryGetLevel(HtmlNode node, out int level) => _nodeLevels.TryGetValue(node, out level);

        private static List<double> CollectBodyFontSizes(HtmlNode root)
        {
            var sizes = new List<double>();

            foreach (var node in root.DescendantsAndSelf())
            {
                if (ShouldSkip(node) || node.NodeType != HtmlNodeType.Element)
                {
                    continue;
                }

                var text = ExtractCandidateText(node);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (TryExtractFontSize(node, out var fontSizePt))
                {
                    sizes.Add(Math.Round(fontSizePt, 2));
                }
            }

            return sizes;
        }

        private static List<HeadingCandidate> CollectBaseCandidates(HtmlNode root)
        {
            var candidates = new List<HeadingCandidate>();

            foreach (var node in root.Descendants())
            {
                if (!IsCandidateBlock(node))
                {
                    continue;
                }

                if (HasBlockedAncestor(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is inside a blocked container.");
                    continue;
                }

                if (IsStructuralContainer(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a structural container.");
                    continue;
                }

                if (node.SelectSingleNode(".//br") is not null)
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is not a single visual line.");
                    continue;
                }

                if (TryConvertOfficeListParagraph(node, 0, out _))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is an Office list paragraph.");
                    continue;
                }

                if (HasQuoteStyle(node) || IsCodeBlock(node))
                {
                    AppLogger.Debug($"Heading candidate rejected: node '{node.Name}' is a quote or code block.");
                    continue;
                }

                var text = ExtractCandidateText(node);
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

                if (IsDisallowedHeadingTerminator(text[^1]))
                {
                    AppLogger.Debug($"Heading candidate rejected: '{text}' ends with disallowed punctuation.");
                    continue;
                }

                if (!TryExtractFontSize(node, out var fontSizePt))
                {
                    AppLogger.Debug($"Heading candidate rejected: '{text}' has no extractable font-size.");
                    continue;
                }

                candidates.Add(new HeadingCandidate(node, text, Math.Round(fontSizePt, 2), IsMostlyBold(node)));
            }

            return candidates;
        }
    }

    private sealed record HeadingCandidate(HtmlNode Node, string Text, double FontSizePt, bool IsBold);
}
