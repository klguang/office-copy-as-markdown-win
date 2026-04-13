# Heading Rules

Chinese version: [heading-rules.zh-CN.md](heading-rules.zh-CN.md)

This document defines the final heading-detection rules used by `office-copy-as-markdown`.

## Overall Flow

Heading detection happens in two steps:

1. First, check whether the current selection already contains semantic headings.
2. Only when there are no semantic headings do candidate heading inference rules apply.

If semantic headings already exist in the current selection, style-based inference is disabled for the entire fragment and only semantic headings are emitted.

## Semantic Heading Priority Rules

Any of the following counts as a semantic heading:

- the HTML node is `h1` through `h6`
- the node style contains `mso-style-name: Heading N`
- the node class name contains `HeadingN`

Mapping rules:

- `Heading 1` or `h1` -> `#`
- `Heading 2` or `h2` -> `##`
- `Heading 3` or `h3` -> `###`
- continue the same way through `Heading 6`

## Base Candidate Heading Rules

Base candidate heading filtering only starts when the current selection does not contain semantic headings.

A base candidate heading must satisfy all of the following:

- occupies a line by itself
- is not a list item
- is not a table cell
- is not a quote
- is not a code block
- has non-empty text
- has no more than `30` characters
- has an extractable valid font size
- does not end with `。 . ， , ； ;`

## Valid Candidate Heading Rules

Base candidates do not become Markdown headings immediately. They must also satisfy either the font-size advantage rule or the "bold and not smaller than body text" rule.

Body-text baseline:

- collect extractable font sizes from the current selection
- use the most common font size as the body-text baseline

Only base candidates that satisfy either of the following become valid candidate headings:

- 1. the font size is at least `2pt` larger than body text, or at least `15%` larger
- 2. the text is bold, and the font size is not smaller than the body-text baseline

Additional notes:

- bold only affects whether a candidate becomes valid; it does not determine the heading level by itself
- the bold rule only applies within the base-candidate set and does not bypass the earlier filters

## Heading Level Mapping Rules

Heading level mapping only applies to valid candidate headings, and only when the current selection does not contain semantic headings.

The maximum supported inferred heading level is controlled by a separate configuration value. The current default is `4`.
When the number of valid font-size tiers is below that limit, the starting mapping level is also controlled by a separate configuration value. The current default is `2`, which means mapping starts from `##`.

Additional notes:

- valid candidate headings are still mapped to Markdown levels only by font-size tier
- bold does not override semantic heading priority and does not introduce a separate "bold heading tier"
- bold candidates with the same font size as body text can still enter the lowest valid font-size tier
- if the number of valid font-size tiers is below the configured maximum level, mapping starts from the configured sparse-mapping start level; the current default behavior starts from `##`

Rules:

1. take the font sizes of all valid candidate headings
2. deduplicate them
3. sort them from largest to smallest
4. if the number of font-size tiers is below the configured maximum supported level, start mapping from the sparse-mapping start level; otherwise start from `#`
5. map the sorted tiers to Markdown heading levels

Mapping:

- first tier -> `#`
- second tier -> `##`
- third tier -> `###`
- fourth tier -> `####`
- fifth tier and below -> no heading output

## Font Size Extraction Rules

Font size extraction order:

1. first read `font-size` from the block node's own `style`
2. if the block node does not contain a font size, continue searching the main text child nodes for `font-size`

Supported units:

- `pt`
- `px`

All values are normalized to `pt` before comparison.

## Typical Examples

### Semantic Headings Present

If the current selection contains:

- `Heading 1`
- a normal short sentence in a large font

then only `Heading 1` is emitted as `#`, and the large-font short sentence does not participate in candidate heading inference.

### No Semantic Headings, Tiered By Font Size

If the current selection contains no semantic headings and the statistics are:

- most common body font size: `12pt`
- candidate heading sizes: `18pt`, `16pt`, `14pt`, `12pt + bold`

then the output is:

- `18pt` -> `#`
- `16pt` -> `##`
- `14pt` -> `###`
- `12pt + bold` -> `####`

### No Semantic Headings, But Fewer Tiers Than The Maximum Supported Level

If the current selection contains no semantic headings and the statistics are:

- most common body font size: `12pt`
- candidate heading size: `12pt + bold`
- maximum supported level config: `4`
- sparse-mapping start level config: `2`

then the output is:

- `12pt + bold` -> `##`

### No Semantic Headings, But Not A Valid Candidate Heading

If the current selection contains no semantic headings and the body-text baseline is `12pt`, then none of the following become valid candidate headings:

- `12pt` but not bold
- `11pt + bold`

### Cases That Never Enter Candidate Heading Inference

Even when they occupy a whole line, the following do not participate in heading inference:

- text longer than `30` characters
- text ending with `。 . ， , ； ;`
- list items
- table cells
- text inside blockquotes
- text inside code blocks

Even if those items are bold, bold does not allow them to bypass the base-candidate filters.

## Known Limitations

- the statistics only cover the current selection and do not scan the whole document
- the candidate heading inference level limit only applies when there are no semantic headings; once semantic headings exist, the entire fragment still outputs only semantic headings
- current candidate heading inference maps at most to the Markdown heading levels allowed by the configured maximum supported level; the current default is `4`
- when the number of font-size tiers is smaller than the configured maximum supported level, inferred results start from the configured sparse-mapping start level; the current default behavior does not consume `#` and starts from `##`
- if the current fragment contains only headings and not enough body-text samples, the body-text baseline may be unstable
- the bold rule only affects valid-candidate detection; list items, table cells, quotes, code blocks, very long text, and text ending in blocked punctuation still cannot become headings just because they are bold
- OneNote and Word do not emit identical HTML structures, so complex styling may still require manual cleanup
