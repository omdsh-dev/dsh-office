// xlsx round-trip tests: write → read → edit → read, through the tool
// execute path exactly as the harness would call it.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import { registerExcelTools } from '../src/excel.js'

let dir: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-excel-'))
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

function collect(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerExcelTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

const noExec = {} as never

describe('xlsx tools', () => {
  it('xlsx_write then xlsx_read round-trips data', async () => {
    const { tools } = collect()
    const file = join(dir, 'book.xlsx')

    const written = await tools['xlsx_write']!.execute({
      file_path: file,
      data: [['Name', 'Score'], ['Alice', 92], ['Bob', 88]],
      sheet_name: 'Results',
      header_bold: true,
    }, noExec) as { content: string }
    expect(written.content).toContain('3 rows')
    expect(existsSync(file)).toBe(true)

    // list sheets
    const listed = await tools['xlsx_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(listed.content).toContain('Results')
    expect(listed.content).toContain('3 rows')

    // read sheet as markdown
    const read = await tools['xlsx_read']!.execute({ file_path: file, sheet: 'Results' }, noExec) as { content: string }
    expect(read.content).toContain('Name')
    expect(read.content).toContain('Alice')
    expect(read.content).toContain('92')
  })

  it('xlsx_write supports formula cells', async () => {
    const { tools } = collect()
    const file = join(dir, 'formula.xlsx')
    await tools['xlsx_write']!.execute({
      file_path: file,
      data: [[1, 2], [{ formula: 'SUM(A1:B1)' }]],
    }, noExec)

    const read = await tools['xlsx_read']!.execute({ file_path: file, sheet: 'Sheet1' }, noExec) as { content: string }
    // formula text is surfaced; cached result may be present
    expect(read.content).toContain('SUM(A1:B1)')
  })

  it('xlsx_edit adds a sheet, updates cells, and appends rows', async () => {
    const { tools } = collect()
    const file = join(dir, 'edit.xlsx')
    await tools['xlsx_write']!.execute({
      file_path: file,
      data: [['A', 'B'], [1, 2]],
    }, noExec)

    const edited = await tools['xlsx_edit']!.execute({
      file_path: file,
      operations: [
        { action: 'add_sheet', name: 'Notes' },
        { action: 'update_cells', sheet: 'Sheet1', cells: [{ cell: 'A3', value: 3 }] },
        { action: 'append_rows', sheet: 'Sheet1', rows: [['C', 4]] },
      ],
    }, noExec) as { content: string }
    expect(edited.content).toContain('added sheet "Notes"')
    expect(edited.content).toContain('updated 1 cell(s)')
    expect(edited.content).toContain('appended 1 row(s)')

    const read = await tools['xlsx_read']!.execute({ file_path: file, sheet: 'Sheet1' }, noExec) as { content: string }
    expect(read.content).toContain('3')
    expect(read.content).toContain('C')
  })

  it('xlsx_read truncates with a continuation hint when max_rows is exceeded', async () => {
    const { tools } = collect()
    const file = join(dir, 'wide.xlsx')
    const rows = [['H'], ...Array.from({ length: 40 }, (_, i) => [`r${i + 1}`])]
    await tools['xlsx_write']!.execute({ file_path: file, data: rows }, noExec)

    const limited = await tools['xlsx_read']!.execute({ file_path: file, sheet: 'Sheet1', max_rows: 10 }, noExec) as { content: string }
    expect(limited.content).toContain('r1')
    expect(limited.content).not.toContain('r35')
    expect(limited.content).toMatch(/Continue with range_start: "A11"/)

    // Follow the hint to read the next chunk.
    const next = await tools['xlsx_read']!.execute({ file_path: file, sheet: 'Sheet1', range_start: 'A11', max_rows: 10 }, noExec) as { content: string }
    expect(next.content).toContain('| r11 |')
    expect(next.content).not.toContain('| r1 |')
  })

  it('xlsx_read on a missing file throws with a clear message', async () => {
    const { tools } = collect()
    await expect(
      tools['xlsx_read']!.execute({ file_path: join(dir, 'nope.xlsx') }, noExec),
    ).rejects.toThrow('File not found')
  })

  it('xlsx_read on a missing sheet reports available sheets', async () => {
    const { tools } = collect()
    const file = join(dir, 'sheets.xlsx')
    await tools['xlsx_write']!.execute({ file_path: file, data: [['x']] }, noExec)
    await expect(
      tools['xlsx_read']!.execute({ file_path: file, sheet: 'Nope' }, noExec),
    ).rejects.toThrow(/not found.*Sheet1/)
  })
})
