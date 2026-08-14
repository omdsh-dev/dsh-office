# @huiliyi37/dsh-office

English | [中文](README.zh.md)

Office document tools for [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) (`dsh`): generate, read, and edit spreadsheets (`.xlsx`), PDFs, and presentations (`.pptx`).

Ported from the office plugins of the [Tianshu](https://github.com/Tianshu-Tui) terminal coding agent (Apache-2.0 licensed upstream), adapted to the dsh cordis tool model.

## Tools

| Tool | What it does |
| --- | --- |
| `xlsx_read` | List sheets of a `.xlsx`, or read one sheet as a markdown table (range-limited for large files; formula text preserved) |
| `xlsx_write` | Write a 2D array to a new `.xlsx` (formula cells, header bold, column widths, number formats) |
| `xlsx_edit` | Edit an existing `.xlsx`: add sheets, update cells (value or formula), append rows |
| `pdf_create` | Generate a real PDF with headings, paragraphs, tables, lists, code blocks and footer page numbers; CJK text renders via an auto-detected system font |
| `pdf_read` | Extract text from a PDF for reading into context |
| `pptx_create` | Generate a `.pptx` deck from slide definitions (title / section / content / two-column / image / table / chart), with optional theme and speaker notes |
| `pptx_read` | Extract slide text as markdown, optionally including speaker notes |

## Install

Install into a dsh profile with the official plugin command (this package
declares a `dsh.bundle` manifest, so `dsh plugin` activates it as a bundle
layer):

```sh
dsh plugin --profile <name> add @huiliyi37/dsh-office
dsh --profile <name>
```

The first `dsh plugin` call initializes the profile (`@deepseek-ai/dsh-base`
is its first bundle) and appends this package to the profile's `bundles` list.
After that, launching the profile loads the plugin and registers all seven
tools automatically.

Manual alternative — install the package anywhere Node resolution can find it
and reference it from your own `cordis.patch.yml`:

```sh
npm install @huiliyi37/dsh-office
```

```yaml
# cordis.patch.yml
- insert:
    - id: dsh-office
      name: '@huiliyi37/dsh-office'
```

## Usage examples

```jsonc
// xlsx_write — create a workbook
{ "file_path": "report.xlsx", "data": [["Name", "Score"], ["Alice", 92]], "header_bold": true }

// pdf_create — document with a heading, table and list
{
  "destination_path": "doc.pdf",
  "title": "Quarterly Report",
  "content": [
    { "type": "heading", "text": "Summary" },
    { "type": "table", "headers": ["Region", "Revenue"], "rows": [["APAC", "120"]] },
    { "type": "list", "items": ["Alpha", "Beta"] }
  ],
  "page_numbers": true
}

// pptx_create — a deck with a title slide and a bullet slide
{
  "destination_path": "deck.pptx",
  "slides": [
    { "type": "title", "title": "Roadmap 2026" },
    { "type": "content", "title": "Highlights", "items": ["Plugin runtime", "Office tools"] }
  ]
}
```

## Development

```sh
npm install
npm run build   # tsc → lib/
npm test        # vitest: round-trip tests through the tool execute path
```

## License

Apache License 2.0. Tool logic ported from the Tianshu office plugins
(also Apache-2.0 licensed, copyright Tianshu contributors); see file
headers for per-module provenance.
