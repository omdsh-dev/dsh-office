// xlsx_recalc / xlsx_audit tests: formula quality gates over the tool
// execute path, exactly as the harness would call them.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import ExcelJS from 'exceljs'
import { registerExcelTools } from '../src/excel.js'

let dir: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-audit-'))
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

async function run(tools: Record<string, ToolDefinition>, name: string, params: Record<string, unknown>): Promise<string> {
  const res = await tools[name]!.execute(params, noExec) as { content: string }
  return res.content
}

describe('xlsx_recalc', () => {
  it('evaluates formulas and reports success for a clean workbook', async () => {
    const { tools } = collect()
    const file = join(dir, 'clean.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        ['Item', 'Qty', 'Price'],
        ['A', 10, 2],
        ['B', 5, 4],
        ['Total', { formula: 'SUM(B2:C3)' }, { formula: 'B4*2' }],
      ],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('status: success')
    expect(out).toContain('total_formulas: 2')
  })

  it('catches division by zero with locations', async () => {
    const { tools } = collect()
    const file = join(dir, 'div0.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [1, 0],
        [{ formula: 'A1/B1' }, { formula: 'A1/0' }],
      ],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('status: errors_found')
    expect(out).toContain('#DIV/0!')
    expect(out).toContain('Sheet1!A2')
  })

  it('propagates errors and detects circular references', async () => {
    const { tools } = collect()
    const file = join(dir, 'circ.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [{ formula: 'B1' }, { formula: 'A1+1' }],
      ],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('errors_found')
    expect(out).toContain('#NUM!')
  })

  it('honors IFERROR wrapping (no error reported)', async () => {
    const { tools } = collect()
    const file = join(dir, 'iferr.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [10, 0],
        [{ formula: 'IFERROR(A1/B1, "n/a")' }],
      ],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('status: success')
  })

  it('evaluates cross-sheet references and VLOOKUP', async () => {
    const { tools } = collect()
    const file = join(dir, 'lookup.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [['A', 1], ['B', 2], ['C', 3]],
      sheet_name: 'Data',
    })
    await run(tools, 'xlsx_edit', {
      file_path: file,
      operations: [
        { action: 'add_sheet', name: 'Summary' },
        { action: 'update_cells', sheet: 'Summary', cells: [
          { cell: 'A1', formula: 'Data!B2' },
          { cell: 'A2', formula: 'VLOOKUP("C", Data!A1:B3, 2, 0)' },
        ] },
      ],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('status: success')
    expect(out).toContain('total_formulas: 2')
  })

  it('reports unsupported functions as warnings', async () => {
    const { tools } = collect()
    const file = join(dir, 'unknown.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [[1], [{ formula: 'MYSTERYFN(A1)' }]],
    })

    const out = await run(tools, 'xlsx_recalc', { file_path: file })
    expect(out).toContain('Unsupported functions')
    expect(out).toContain('MYSTERYFN')
  })

  it('fails with a clear message for a missing file', async () => {
    const { tools } = collect()
    await expect(
      tools['xlsx_recalc']!.execute({ file_path: join(dir, 'nope.xlsx') }, noExec),
    ).rejects.toThrow('File not found')
  })
})

describe('xlsx_audit', () => {
  it('flags the array-formula trap aggregate(IF(range,...))', async () => {
    const { tools } = collect()
    const file = join(dir, 'trap.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [1, 2, 3],
        [{ formula: 'MEDIAN(IF(A1:C1>0, A1:C1))' }],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('array_formula_risk')
    expect(out).toContain('Sheet1!A2')
  })

  it('flags aggregation ranges that miss an adjacent data row', async () => {
    const { tools } = collect()
    const file = join(dir, 'gap.xlsx')
    // B1 header, B2 value 5 (missed), B3..B4 summed
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        ['hdr', 'Val'],
        [10, 5],
        [20, 7],
        [{ formula: 'SUM(B3:B4)' }, { formula: 'SUM(A3:A4)' }],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('range_gap')
    // A3:A4 is fine (A2 holds 10 but A1 is a header — suppressed by header rule? no: A2 is number)
    // B3:B4 misses B2 → flagged; A3:A4 misses A2 → also flagged by the heuristic
    expect(out).toContain('missing row 2')
  })

  it('flags self-references and literal division by zero', async () => {
    const { tools } = collect()
    const file = join(dir, 'self.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [1],
        [{ formula: 'A2+1' }, { formula: 'A1/0' }],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('self_reference')
    expect(out).toContain('division_by_zero')
  })

  it('flags hardcoded numbers adjacent to a formula run (possible overwrite)', async () => {
    const { tools } = collect()
    const file = join(dir, 'overwrite.xlsx')
    // B2:B3 formulas, B4 a bare number
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        ['A', 'B'],
        [1, { formula: 'A2*2' }],
        [2, { formula: 'A3*2' }],
        [3, 99],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('possible_overwrite')
  })

  it('flags inconsistent formula structure within a column', async () => {
    const { tools } = collect()
    const file = join(dir, 'inconsistent.xlsx')
    // A2:A6 = A1*2 pattern; A4 uses A1+1 instead
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [1],
        [{ formula: 'A1*2' }],
        [{ formula: 'A2*2' }],
        [{ formula: 'A3+1' }],
        [{ formula: 'A4*2' }],
        [{ formula: 'A5*2' }],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('inconsistent_formula')
  })

  it('detects cached error values saved by other tools', async () => {
    const { tools } = collect()
    const file = join(dir, 'cached.xlsx')
    // hand-build a workbook with a cached error result
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S1')
    ws.getCell('A1').value = { formula: '1/0', result: '#DIV/0!' }
    await wb.xlsx.writeFile(file)

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('cached_error_value')
    expect(out).toContain('S1!A1')
  })

  it('reports clean for a straightforward workbook', async () => {
    const { tools } = collect()
    const file = join(dir, 'ok.xlsx')
    await run(tools, 'xlsx_write', {
      file_path: file,
      data: [
        [1, 2, 3],
        [{ formula: 'SUM(A1:C1)' }],
      ],
    })

    const out = await run(tools, 'xlsx_audit', { file_path: file })
    expect(out).toContain('status: clean')
  })
})
