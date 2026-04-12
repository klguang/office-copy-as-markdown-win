using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace Html2Markdown;

internal static partial class MarkdownConverter
{
    private static string ConvertInlines(HtmlNode node, ConversionContext context)
    {
        if (node.Name.Equals("pre", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertCodeBlock(node);
        }

        return ConvertNodesAsInline(node.ChildNodes, context);
    }

    private static string ConvertNodesAsInline(IEnumerable<HtmlNode> nodes, ConversionContext context)
    {
        var builder = new StringBuilder();

        foreach (var node in nodes)
        {
            if (ShouldSkip(node))
            {
                continue;
            }

            if (node.NodeType == HtmlNodeType.Text)
            {
                builder.Append(NormalizeInlineText(WebUtility.HtmlDecode(node.InnerText)));
                continue;
            }

            builder.Append(ConvertInlineElement(node, context));
        }

        return NormalizeInlineResult(builder.ToString());
    }

    private static string ConvertInlineElement(HtmlNode node, ConversionContext context)
    {
        if (node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
        {
            return "  \n";
        }

        if (node.Name.Equals("a", StringComparison.OrdinalIgnoreCase))
        {
            var text = ConvertNodesAsInline(node.ChildNodes, context).Trim();
            var href = node.GetAttributeValue("href", string.Empty).Trim();
            return string.IsNullOrWhiteSpace(href) ? text : $"[{text}]({href})";
        }

        if (node.Name.Equals("img", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertImage(node, context);
        }

        if (node.Name.Equals("code", StringComparison.OrdinalIgnoreCase) || IsInlineCode(node))
        {
            var inlineCode = ConvertNodesAsInline(node.ChildNodes, context).Trim();
            return string.IsNullOrWhiteSpace(inlineCode) ? string.Empty : WrapInlineCode(inlineCode);
        }

        if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase) || node.Name.Equals("b", StringComparison.OrdinalIgnoreCase))
        {
            return Wrap("**", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name.Equals("em", StringComparison.OrdinalIgnoreCase) || node.Name.Equals("i", StringComparison.OrdinalIgnoreCase))
        {
            return Wrap("*", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name is "del" or "s" or "strike" || HasStrikeStyle(node))
        {
            return Wrap("~~", ConvertNodesAsInline(node.ChildNodes, context));
        }

        if (node.Name.Equals("input", StringComparison.OrdinalIgnoreCase))
        {
            return ConvertCheckbox(node);
        }

        var content = ConvertNodesAsInline(node.ChildNodes, context);

        if (HasBoldStyle(node))
        {
            content = Wrap("**", content);
        }

        if (HasItalicStyle(node))
        {
            content = Wrap("*", content);
        }

        if (HasStrikeStyle(node))
        {
            content = Wrap("~~", content);
        }

        if (IsInlineCode(node))
        {
            content = WrapInlineCode(content);
        }

        return content;
    }

    private static string ConvertImage(HtmlNode node, ConversionContext context)
    {
        var src = node.GetAttributeValue("src", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(src) && context.CanUseFallbackImage)
        {
            src = context.UseFallbackImage();
        }

        if (string.IsNullOrWhiteSpace(src))
        {
            return string.Empty;
        }

        var alt = node.GetAttributeValue("alt", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(alt))
        {
            alt = "image";
        }

        return $"![{alt}]({src})";
    }

    private static string Wrap(string marker, string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        var leadingWhitespace = content.Length - content.TrimStart().Length;
        var trailingWhitespace = content.Length - content.TrimEnd().Length;
        var trimmed = content.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        return $"{new string(' ', leadingWhitespace)}{marker}{trimmed}{marker}{new string(' ', trailingWhitespace)}";
    }

    private static string WrapInlineCode(string content)
    {
        var trimmed = content.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        return $"`{trimmed.Replace("`", "\\`", StringComparison.Ordinal)}`";
    }

    private static string NormalizeInlineText(string text)
    {
        var normalized = text
            .Replace('\u00A0', ' ')
            .Replace('\u200B', ' ');

        return MarkdownLineSyntax.NormalizeWhitespace(normalized);
    }

    private static string NormalizeInlineResult(string text)
    {
        var normalized = Regex.Replace(text, @"[ \t]+(\r?\n)", "$1").Trim();
        while (true)
        {
            var previous = normalized;
            normalized = AdjacentBoldSegmentsRegex.Replace(normalized, match =>
                $"**{match.Groups["left"].Value}{match.Groups["middle"].Value}{match.Groups["right"].Value}**");
            normalized = AdjacentBoldRunsRegex.Replace(normalized, match =>
            {
                var left = match.Groups["left"].Value;
                var right = match.Groups["right"].Value;
                return $"**{left}{ResolveAdjacentBoldJoiner(left, right)}{right}**";
            });

            if (string.Equals(previous, normalized, StringComparison.Ordinal))
            {
                break;
            }
        }

        normalized = BoldLabelSeparatedColonRegex.Replace(normalized, match =>
        {
            var label = match.Groups["label"].Value;
            var colon = match.Groups["colon"].Value;
            var next = match.Groups["next"].Value;
            if (string.IsNullOrEmpty(label) || string.IsNullOrEmpty(colon) || string.IsNullOrEmpty(next))
            {
                return match.Value;
            }

            if (!ShouldInsertSpaceAfterBoldLabel(colon[0], next[0]))
            {
                return match.Value;
            }

            return $"**{label}{colon}** {next}";
        });

        normalized = BoldLabelFollowedByTextRegex.Replace(normalized, match =>
        {
            var label = match.Groups["label"].Value;
            var next = match.Groups["next"].Value;
            if (string.IsNullOrEmpty(label) || string.IsNullOrEmpty(next))
            {
                return match.Value;
            }

            if (!ShouldInsertSpaceAfterBoldLabel(label[^1], next[0]))
            {
                return match.Value;
            }

            return $"**{label}** {next}";
        });

        normalized = TextFollowedByBoldLabelRegex.Replace(normalized, match =>
        {
            var previous = match.Groups["prev"].Value;
            var label = match.Groups["label"].Value;
            if (string.IsNullOrEmpty(previous) || string.IsNullOrEmpty(label))
            {
                return match.Value;
            }

            if (!ShouldInsertSpaceBeforeBoldLabel(previous[0], label[0]))
            {
                return match.Value;
            }

            return $"{previous} **{label}**";
        });

        normalized = TightPlusBetweenWordsRegex.Replace(normalized, "${left} + ${right}");

        return normalized;
    }

    private static string NormalizeHeadingText(string text)
    {
        var normalized = text.Trim();

        normalized = normalized
            .Replace("**", string.Empty, StringComparison.Ordinal)
            .Replace("__", string.Empty, StringComparison.Ordinal);

        normalized = NormalizeInlineText(normalized).Trim();
        normalized = NumberedHeadingRegex.Replace(normalized, "${prefix} ${rest}");
        return normalized;
    }

    private static string ResolveAdjacentBoldJoiner(string left, string right)
    {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
        {
            return string.Empty;
        }

        var leftChar = left[^1];
        var rightChar = right[0];

        if (char.IsWhiteSpace(leftChar) || char.IsWhiteSpace(rightChar))
        {
            return string.Empty;
        }

        if (IsOpeningPunctuation(leftChar) || IsClosingPunctuation(rightChar))
        {
            return string.Empty;
        }

        if (IsSeparatorPunctuation(leftChar) || IsSeparatorPunctuation(rightChar))
        {
            return string.Empty;
        }

        if (HasUnclosedOpeningPunctuation(left))
        {
            return string.Empty;
        }

        if ((IsCjk(leftChar) && IsLatinOrDigit(rightChar)) ||
            (IsLatinOrDigit(leftChar) && IsCjk(rightChar)) ||
            (char.IsLetterOrDigit(leftChar) && char.IsLetterOrDigit(rightChar)))
        {
            return " ";
        }

        return string.Empty;
    }

    private static bool ShouldInsertSpaceAfterBoldLabel(char labelEnd, char nextChar)
    {
        if (labelEnd is not (':' or '\uFF1A'))
        {
            return false;
        }

        if (char.IsWhiteSpace(nextChar) ||
            IsOpeningPunctuation(nextChar) ||
            IsClosingPunctuation(nextChar) ||
            IsSeparatorPunctuation(nextChar))
        {
            return false;
        }

        return IsCjk(nextChar) || IsLatinOrDigit(nextChar);
    }

    private static bool ShouldInsertSpaceBeforeBoldLabel(char previousChar, char labelStart)
    {
        if (!IsCjk(previousChar) && !IsLatinOrDigit(previousChar))
        {
            return false;
        }

        if (!IsCjk(labelStart) && !IsLatinOrDigit(labelStart))
        {
            return false;
        }

        if (IsOpeningPunctuation(previousChar) || IsSeparatorPunctuation(previousChar))
        {
            return false;
        }

        return true;
    }

    private static bool HasUnclosedOpeningPunctuation(string text)
    {
        var depth = 0;
        foreach (var ch in text)
        {
            if (IsOpeningPunctuation(ch))
            {
                depth++;
                continue;
            }

            if (IsClosingPunctuation(ch) && depth > 0)
            {
                depth--;
            }
        }

        return depth > 0;
    }

    private static bool IsCjk(char value) => value is >= '\u4E00' and <= '\u9FFF';

    private static bool IsLatinOrDigit(char value) =>
        (value is >= 'A' and <= 'Z') ||
        (value is >= 'a' and <= 'z') ||
        char.IsDigit(value);

    private static bool IsOpeningPunctuation(char value) => value is '(' or '[' or '{' or '<' or '\uFF08' or '\u300A' or '\u201C' or '"' or '\'';

    private static bool IsClosingPunctuation(char value) => value is ')' or ']' or '}' or '>' or '\uFF09' or '\u300B' or '\u201D' or '"' or '\'';

    private static bool IsSeparatorPunctuation(char value) => value is ':' or '\uFF1A' or '/' or '\\' or '-' or '\u2014' or '+' or '\uFF0B';
}
