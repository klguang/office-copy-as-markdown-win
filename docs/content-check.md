# Content Completeness And Backfill Rules

Chinese version: [content-check.zh-CN.md](content-check.zh-CN.md)

This document defines the rules used to verify Markdown completeness and backfill missing content.

## Goals

1. The final output must be Markdown.
2. The visible text in Markdown should cover the source plain text as completely as possible.
3. When structured conversion conflicts with content completeness, content completeness wins.
4. Raw HTML must not be emitted just to recover missing content.

## Baseline Rules

1. Use the source plain text as the completeness baseline.
2. Use the converted Markdown as the candidate result to validate.
3. Comparison only considers visible text; layout does not need to match Office exactly.
4. Headings may only come from the original structure; the backfill stage must not introduce new headings.

## Source Text Splitting Rules

1. Normalize line endings.
2. Split by line.
3. Skip blank lines.
4. Classify each line, in order, as one of:
   - task list
   - ordered list
   - unordered list
   - normal paragraph
5. The detected line type must be preserved during backfill.

## Markdown Comparison Rules

1. Images are compared by visible alternative text.
2. Links are compared by link text.
3. Bold, italic, strikethrough, and code markup do not participate in comparison.
4. Heading, quote, and list prefixes are stripped before comparison.
5. Empty result lines do not participate in comparison.

## Normalization Rules

1. Collapse consecutive whitespace into a single space.
2. Treat `Tab` as a space.
3. Normalize `→` to `->`.
4. Normalize full-width `＋` to `+`.
5. Optional spaces between Chinese and English text are not treated as differences.
6. Extra spaces between Chinese characters are not treated as differences.
7. Formatting spaces around `/`, `:`, `：`, and `+` are not treated as differences.

## Table Rules

1. Compare Markdown tables and plain-text tables by cell content.
2. Table separator rows do not participate in comparison.
3. Matching cell text is sufficient to count as covered.
4. Different table syntax or plain-text separators must not be treated as missing content by themselves.

## Completeness Decision Rules

1. Use ordered matching.
2. For each source line, search only for the first match that appears after the current Markdown position.
3. Any unmatched source text counts as missing content.
4. The result is complete only when the number of missing items is `0`.

## Backfill Rules

1. Missing content must be backfilled into Markdown.
2. Preserve the original line type during backfill:
   - task lists stay task lists
   - ordered lists stay ordered lists
   - unordered lists stay unordered lists
   - normal text becomes paragraphs
3. Prefer inserting after the previous matched anchor.
4. If there is no previous anchor, insert before the next matched anchor.
5. If there are no anchors on either side, append to the end of the document.
6. Minimal blank-line adjustment is allowed after backfill, but do not create unnecessary runs of empty blocks.

## Fallback Rules

1. If structured Markdown fails validation, run backfill first.
2. If the result is still incomplete after backfill, fall back to conservative Markdown.
3. Conservative Markdown only requires complete content, not optimal structure.
4. Conservative Markdown must still pass completeness validation again.

## Forbidden Rules

1. Content must not be silently dropped when validation fails.
2. Normal text must not be upgraded to headings during backfill.
3. Existing content must not be duplicated just to pass validation.
4. Representation differences for tables, links, images, or mixed Chinese and English text must not be misclassified as missing content.

## One-Sentence Principle

Generate the best Markdown you can first, then compare it line by line against the source plain text; anything not covered by the Markdown must be backfilled in Markdown form until the result is complete.
