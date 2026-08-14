// xlsx read/write/edit tools, ported from the Tianshu office-excel plugin
// (Apache-2.0 licensed upstream) to the DeepSeek Harness cordis tool model.

import { existsSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { defineTool, type JsonValue } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { textOutput } from './text-output.js'

// ── shared helpers ─────────────────────────────────────────────────

function colToIndex(col: string | undefined): number {
  if (typeof col !== 'string' || col.length === 0) return -1
  let result = 0
  for (const ch of col.toUpperCase()) {
    result = result * 26 + (ch.charCodeAt(0) - 64)
  }
  return result
}

function padRight(str: string, len: number): string {
  return str + ' '.repeat(Math.max(0, len - str.length))
}

interface CellValueInput {
  formula?: string
  result?: JsonValue
  value?: JsonValue
}

// Cell values may be plain scalars or { formula: 'SUM(A1:A9)' } objects.
function applyCellValue(cell: ExcelJS.Cell, val: unknown): void {
  if (val && typeof val === 'object' && !Array.isArray(val) && 'formula' in val) {
    const formula = String((val as { formula: string }).formula).replace(/^=/, '')
    const result = (val as { result?: JsonValue }).result
    cell.value = result !== undefined && result !== null
      ? { formula, result: result as number | string }
      : { formula }
  } else {
    cell.value = (val ?? null) as never
  }
}

function addRowValues(ws: ExcelJS.Worksheet, rowData: JsonValue[]): void {
  const row = ws.addRow([])
  rowData.forEach((val, i) => applyCellValue(row.getCell(i + 1), val))
}

interface StyleParams {
  header_bold?: boolean
  column_widths?: JsonValue[]
  number_formats?: Record<string, JsonValue> | null
}

// Basic styling: header_bold, column_widths, number_formats ({ B: '#,##0.00' }).
function applyStyles(ws: ExcelJS.Worksheet, params: StyleParams): void {
  if (params?.header_bold && ws.rowCount >= 1) {
    ws.getRow(1).font = { bold: true }
  }
  const widths = params?.column_widths
  if (Array.isArray(widths)) {
    widths.forEach((w, i) => {
      if (typeof w === 'number' && w > 0) ws.getColumn(i + 1).width = w
    })
  }
  const formats = params?.number_formats
  if (formats && typeof formats === 'object') {
    for (const [col, fmt] of Object.entries(formats)) {
      ws.getColumn(colToIndex(col)).numFmt = String(fmt)
    }
  }
}

// ── xlsx_read ──────────────────────────────────────────────────────

async function xlsxRead(params: {
  file_path: string
  sheet?: string
  range_start?: string
  range_end?: string
  max_rows?: number
}): Promise<string> {
  const filePath = params.file_path
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`)
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const sheetName = params.sheet
  const rangeStart = params.range_start // e.g. "A1"
  const rangeEnd = params.range_end     // e.g. "D20"

  // List sheets mode
  if (!sheetName) {
    const sheets = workbook.worksheets.map(ws => ({
      name: ws.name,
      rows: ws.rowCount,
      cols: ws.columnCount,
    }))
    return [
      `Workbook: ${filePath}`,
      `Sheets: ${sheets.length}`,
      '',
      ...sheets.map(s => `  ${s.name} — ${s.rows} rows × ${s.cols} cols`),
      '',
      'Use sheet parameter to read a specific sheet. Add range_start/range_end for partial read.',
    ].join('\n')
  }

  // Read sheet mode
  const ws = workbook.getWorksheet(sheetName)
  if (!ws) {
    const available = workbook.worksheets.map(w => w.name).join(', ')
    throw new Error(`Sheet "${sheetName}" not found. Available: ${available}`)
  }

  // Determine range — ExcelJS cell addresses: parse manually
  let startRow = 1, startCol = 1
  let endRow = ws.rowCount, endCol = ws.columnCount

  if (rangeStart) {
    const match = rangeStart.match(/^([A-Z]+)(\d+)$/i)
    if (match) {
      startCol = colToIndex(match[1] ?? '')
      startRow = parseInt(match[2] ?? '0', 10)
    }
  }
  if (rangeEnd) {
    const match = rangeEnd.match(/^([A-Z]+)(\d+)$/i)
    if (match) {
      endCol = colToIndex(match[1] ?? '')
      endRow = parseInt(match[2] ?? '0', 10)
    }
  }

  // Clamp
  endRow = Math.min(endRow, ws.rowCount)
  endCol = Math.min(endCol, ws.columnCount || 26)

  // Read cells into markdown table
  const rows: string[][] = []
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r)
    const cells: string[] = []
    for (let c = startCol; c <= endCol; c++) {
      const cell = row.getCell(c)
      const val = cell.value
      if (val && typeof val === 'object' && 'formula' in val) {
        // Formula cell: show cached result when present, always keep formula text
        const text = `=${String((val as { formula: string }).formula)}`
        const result = (val as { result?: unknown }).result
        cells.push(result !== undefined && result !== null ? `${String(result)} (${text})` : text)
      } else if (val && typeof val === 'object' && 'result' in val) {
        cells.push(String((val as { result?: unknown }).result ?? ''))
      } else if (val !== null && val !== undefined) {
        cells.push(String(val))
      } else {
        cells.push('')
      }
    }
    rows.push(cells)
  }

  if (rows.length === 0) {
    return `Sheet "${sheetName}" is empty.`
  }

  // Render markdown table, truncated at the page boundary (default 200 rows
  // for context safety). The continuation hint hands the model the exact
  // range_start to read the next chunk.
  const maxRows = Math.min(Math.max(1, params.max_rows ?? 200), 500)
  const displayRows = rows.slice(0, maxRows)
  const colWidths: number[] = []
  for (let c = 0; c < (displayRows[0]?.length || 0); c++) {
    let max = 3
    for (const row of displayRows) {
      max = Math.max(max, (row[c] || '').length)
    }
    colWidths.push(Math.min(max, 40))
  }

  const mdRows = displayRows.map((row, i) => {
    const cells = row.map((cell, ci) => padRight(String(cell).slice(0, 40), colWidths[ci] || 3))
    return '| ' + cells.join(' | ') + ' |'
  })

  // Header separator
  if (mdRows.length > 0) {
    const sep = '|' + colWidths.map(w => '-'.repeat(w + 2)).join('|') + '|'
    mdRows.splice(1, 0, sep)
  }

  const suffix = rows.length > maxRows
    ? `\n\n(Showing ${maxRows} of ${rows.length} rows. Continue with range_start: "A${startRow + maxRows}".)`
    : ''

  return `Sheet "${sheetName}" (${rows.length} rows × ${endCol - startCol + 1} cols):\n\n${mdRows.join('\n')}${suffix}`
}

// ── xlsx_write ─────────────────────────────────────────────────────

async function xlsxWrite(params: {
  file_path: string
  data: JsonValue[][]
  sheet_name?: string
  header_bold?: boolean
  column_widths?: JsonValue[]
  number_formats?: Record<string, JsonValue> | null
}): Promise<string> {
  const filePath = params.file_path || (params as { destination_path?: string }).destination_path
  if (!filePath) {
    throw new Error('Missing file_path parameter')
  }

  const data = params.data
  if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0])) {
    throw new Error('Missing or invalid data: expected 2D array')
  }

  const workbook = new ExcelJS.Workbook()
  const sheetName = params.sheet_name || 'Sheet1'
  const ws = workbook.addWorksheet(sheetName)

  for (const rowData of data) {
    addRowValues(ws, rowData)
  }

  applyStyles(ws, params)

  await workbook.xlsx.writeFile(filePath)

  return `Written ${data.length} rows × ${(data[0] as JsonValue[]).length} cols to ${filePath} (sheet: "${sheetName}")`
}

// ── xlsx_edit ──────────────────────────────────────────────────────

interface EditOperation {
  action: 'add_sheet' | 'update_cells' | 'append_rows'
  name?: string
  sheet?: string
  cells?: Array<{ cell?: string; value?: JsonValue; formula?: string }>
  rows?: JsonValue[][]
}

async function xlsxEdit(params: {
  file_path: string
  output_path?: string
  operations: EditOperation[]
  style_sheet?: string
  header_bold?: boolean
  column_widths?: JsonValue[]
  number_formats?: Record<string, JsonValue> | null
}): Promise<string> {
  const filePath = params.file_path
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`)
  }

  const operations = params.operations
  if (!Array.isArray(operations) || operations.length === 0) {
    throw new Error('Missing or invalid operations: expected non-empty array')
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const applied: string[] = []
  for (const op of operations) {
    switch (op?.action) {
      case 'add_sheet': {
        const name = op.name
        if (!name || typeof name !== 'string') {
          throw new Error('add_sheet requires a name string')
        }
        if (workbook.getWorksheet(name)) {
          applied.push(`sheet "${name}" already exists — skipped`)
        } else {
          workbook.addWorksheet(name)
          applied.push(`added sheet "${name}"`)
        }
        break
      }
      case 'update_cells': {
        const ws = workbook.getWorksheet(op.sheet || '')
        if (!ws) {
          const available = workbook.worksheets.map(w => w.name).join(', ')
          throw new Error(`Sheet "${op.sheet}" not found. Available: ${available}`)
        }
        if (!Array.isArray(op.cells) || op.cells.length === 0) {
          throw new Error('update_cells requires a non-empty cells array')
        }
        for (const c of op.cells) {
          if (!c?.cell) {
            throw new Error('update_cells: each entry needs a cell address (e.g. "B2")')
          }
          applyCellValue(ws.getCell(c.cell), 'formula' in c ? { formula: c.formula } : (c.value ?? null))
        }
        applied.push(`updated ${op.cells.length} cell(s) in "${op.sheet}"`)
        break
      }
      case 'append_rows': {
        const ws = workbook.getWorksheet(op.sheet || '')
        if (!ws) {
          const available = workbook.worksheets.map(w => w.name).join(', ')
          throw new Error(`Sheet "${op.sheet}" not found. Available: ${available}`)
        }
        if (!Array.isArray(op.rows) || op.rows.length === 0 || !Array.isArray(op.rows[0])) {
          throw new Error('append_rows requires a non-empty 2D rows array')
        }
        for (const rowData of op.rows) {
          addRowValues(ws, rowData)
        }
        applied.push(`appended ${op.rows.length} row(s) to "${op.sheet}"`)
        break
      }
      default:
        throw new Error(`Unknown action: ${String(op?.action)}. Supported: add_sheet, update_cells, append_rows`)
    }
  }

  // Optional styles apply to style_sheet (default: first worksheet)
  if (params?.header_bold || params?.column_widths || params?.number_formats) {
    const styleWs = params?.style_sheet
      ? workbook.getWorksheet(params.style_sheet)
      : workbook.worksheets[0]
    if (!styleWs) {
      throw new Error(`Sheet "${params?.style_sheet}" not found for styling`)
    }
    applyStyles(styleWs, params)
    applied.push(`applied styles to "${styleWs.name}"`)
  }

  const outPath = params?.output_path || filePath
  await workbook.xlsx.writeFile(outPath)

  return `Edited ${filePath}${outPath !== filePath ? ` → ${outPath}` : ''}\n${applied.map(a => `  - ${a}`).join('\n')}`
}

// ── registration ───────────────────────────────────────────────────

export function registerExcelTools(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'xlsx_read',
    description: 'Read a .xlsx file: list all sheets, or read a specific sheet as a markdown table. Supports range_start/range_end and max_rows for large files (a continuation range_start hint is appended on truncation). Formula cells show the formula text.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .xlsx file' },
      sheet: { type: 'string', description: 'Sheet name to read (omit to list sheets)' },
      range_start: { type: 'string', description: 'Start cell e.g. "A1"' },
      range_end: { type: 'string', description: 'End cell e.g. "D20"' },
      max_rows: { type: 'integer', description: 'Max rows per read (1-500, default 200); a continuation range_start hint is appended when truncated' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await xlsxRead(args) }),
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'xlsx_write',
    description: 'Write a 2D array to a new .xlsx file. Cell values can be strings, numbers, booleans, or { formula: "SUM(A1:A9)" } objects. Supports header_bold / column_widths / number_formats.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Destination .xlsx file path' },
      data: {
        type: 'array', required: true, items: { type: 'array' },
        description: '2D array of cell values; use { formula: "..." } for formula cells',
      },
      sheet_name: { type: 'string', description: 'Sheet name (default: Sheet1)' },
      header_bold: { type: 'boolean', description: 'Bold the first row' },
      column_widths: { type: 'array', items: { type: 'number' }, description: 'Column widths by position, e.g. [12, 20, 20]' },
      number_formats: { type: 'object', additionalProperties: true, description: 'numFmt per column letter, e.g. { "B": "#,##0.00" }' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await xlsxWrite(args) }),
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'xlsx_edit',
    description: 'Edit an existing .xlsx file: add sheets, update individual cells (value or formula), append rows. Saves back to file_path unless output_path is given.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the existing .xlsx file' },
      output_path: { type: 'string', description: 'Save to a different path instead of overwriting' },
      operations: {
        type: 'array', required: true,
        description: 'Ordered edit operations',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            action: { type: 'string', required: true, enum: ['add_sheet', 'update_cells', 'append_rows'] as const },
            name: { type: 'string', description: 'add_sheet: new sheet name' },
            sheet: { type: 'string', description: 'update_cells/append_rows: target sheet name' },
            cells: {
              type: 'array',
              description: 'update_cells: [{ cell: "B2", value: 42 } or { cell: "B3", formula: "SUM(B1:B2)" }]',
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  cell: { type: 'string', required: true, description: 'Cell address, e.g. "B2"' },
                  value: { type: 'json', description: 'Cell value (scalar or { formula: "..." })' },
                  formula: { type: 'string', description: 'Formula text without leading "=", e.g. "SUM(B1:B2)"' },
                },
              },
            },
            rows: { type: 'array', items: { type: 'array' }, description: 'append_rows: 2D array of cell values' },
          },
        },
      },
      style_sheet: { type: 'string', description: 'Sheet the style options apply to (default: first sheet)' },
      header_bold: { type: 'boolean', description: 'Bold the first row' },
      column_widths: { type: 'array', items: { type: 'number' }, description: 'Column widths by position, e.g. [12, 20, 20]' },
      number_formats: { type: 'object', additionalProperties: true, description: 'numFmt per column letter, e.g. { "B": "#,##0.00" }' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await xlsxEdit(args) }),
    isConcurrencySafe: () => false,
  }))
}
