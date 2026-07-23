# Purpose

Twice a month I have to convert a `.docx` file to BBcode in order to fulfill a function as a non-profit board secretary. It's a bit tedious. This automates it.

Converts a `.docx` file to phpBB BBCode markup, preserving:
- Bold / italic / underline / strikethrough
- Hyperlinks, including both real `w:hyperlink` elements **and** Word "field code" hyperlinks (`{ HYPERLINK "url" }display text{ }`)
- Nested bulleted/numbered lists (via `numPr`, regardless of paragraph style)
- Manual line breaks (`w:br`) and blank spacer paragraphs

# Usage

    python3 docx2bbcode.py input.docx > output.bbcode.txt

# Disclaimer

Claude wrote this; don't ask me how it works.
