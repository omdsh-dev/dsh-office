# @huiliyi37/dsh-office

English | [中文](README.zh.md)

Office document tools for [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) (`dsh`): generate, read, and edit spreadsheets (`.xlsx`), PDFs, presentations (`.pptx`), and Word documents (`.docx`).

Ported from the office plugins of the [Tianshu](https://github.com/Tianshu-Tui) terminal coding agent (Apache-2.0 licensed upstream), adapted to the dsh cordis tool model.

## Tools

| Tool | What it does |
| --- | --- |
| `xlsx_read` | List sheets of a `.xlsx`, or read one sheet as a markdown table (range-limited for large files; formula text preserved) |
| `xlsx_write` | Write a 2D array to a new `.xlsx` (formula cells, header bold, column widths, number formats) |
| `xlsx_edit` | Edit an existing `.xlsx`: add sheets, update cells (value or formula), append rows |
| `xlsx_recalc` | Recalculate every formula with a lightweight pure-TS engine and report error values (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`, `#NUM!`) with locations |
| `xlsx_audit` | Statically audit formula structure: array-formula traps, aggregation ranges missing rows, formulas overwritten by hardcoded values, inconsistent formulas in a column, self-references, division by zero |
| `pdf_create` | Generate a real PDF with headings, paragraphs, tables, lists, code blocks and footer page numbers; CJK text renders via an auto-detected system font |
| `pdf_read` | Extract text from a PDF for reading into context |
| `pdf_merge` | Merge multiple PDFs into one, in order |
| `pdf_split` | Split a PDF into single-page files, or extract pages by spec (`"1,3,5-7"`) |
| `pptx_create` | Generate a `.pptx` deck from slide definitions (title / section / content / two-column / image / table / chart), with optional theme and speaker notes |
| `pptx_read` | Extract slide text as markdown, optionally including speaker notes; `include` adds structure (summary/layouts/images/tables: shape names & cm positions, image targets, table dims) |
| `pptx_edit` | Find/replace text inside an existing `.pptx` (`<a:t>` surgery), preserving all layout and styling |
| `docx_create` | Generate a real `.docx` Word document from content blocks (heading / paragraph / table / code / list), with optional striped tables, page background color, and diagonal text watermark |
| `docx_read` | Extract text from a `.docx` for reading into context |

## Install & load

> **Compatibility**: requires `dsh` ≥ `0.1.0-rc.5` (bundles `@deepseek-ai/dsh-tools` ≥ `0.1.0-rc.5`). Installing on an older core whose `dsh-tools` is a `0.0.1` release installs a second `dsh-tools` copy and crashes every tool call with `Cannot read properties of undefined (reading 'prepare')`.

### Full load (all 14 tools)

```sh
dsh plugin --profile <name> add @huiliyi37/dsh-office
dsh --profile <name>
```

The first `dsh plugin` call initializes the profile (`@deepseek-ai/dsh-base`
is its first bundle) and appends this package to the profile's `bundles` list.
Launching the profile then registers all tools automatically.

### Load only the families you need

Pass a `config` row in your profile's `cordis.patch.yml` to enable/disable
per family. Omitted families stay enabled; set a family to `false` to
exclude it (useful to keep the tool surface small):

```yaml
# cordis.patch.yml
- insert:
    - id: dsh-office
      name: '@huiliyi37/dsh-office'
      config:
        enable:
          xlsx: false      # skip spreadsheet tools
          pdf: true        # keep PDF tools
          ppt: false       # skip presentations
          docx: true       # keep Word tools
```

Families: `xlsx` (read/write/edit/recalc/audit), `pdf` (create/read/merge/split),
`ppt` (create/read/edit), `docx` (create/read).

### Uninstall

```sh
dsh plugin --profile <name> remove @huiliyi37/dsh-office
```

### Manual alternative

Install the package anywhere Node resolution can find it and reference it
from your own `cordis.patch.yml` (same `config.enable` switches apply):

```sh
npm install @huiliyi37/dsh-office
```

```yaml
# cordis.patch.yml
- insert:
    - id: dsh-office
      name: '@huiliyi37/dsh-office'
```

## Skill

The package ships a usage skill (`skills/SKILL.md`, anthropics-compatible
format) that teaches the model large-file pagination and generation
discipline. Install it into a skill discovery root:

```sh
mkdir -p ~/.dsh/skills && cp -r node_modules/@huiliyi37/dsh-office/skills/dsh-office ~/.dsh/skills/
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
