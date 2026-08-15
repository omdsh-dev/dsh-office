// pptx_read structured query (include) + pptx_edit text surgery tests,
// through the tool execute path.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import { registerPptTools } from '../src/ppt.js'

let dir: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-ppt-'))
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

function collect(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerPptTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

const noExec = {} as never

async function run(tools: Record<string, ToolDefinition>, name: string, params: Record<string, unknown>): Promise<string> {
  const res = await tools[name]!.execute(params, noExec) as { content: string }
  return res.content
}

function deckPath(name: string): string {
  return join(dir, name)
}

/** Build a small deck: title slide + content slide with bullet items + a table slide. */
async function makeDeck(file: string): Promise<Record<string, ToolDefinition>> {
  const { tools } = collect()
  await run(tools, 'pptx_create', {
    destination_path: file,
    title: 'Demo Deck',
    slides: [
      { type: 'title', title: '季度汇报', body: '2026 Q3', notes: '开场白' },
      { type: 'content', title: '业绩要点', items: ['营收增长 12%', '成本下降 5%'] },
      { type: 'table', title: '数据明细', headers: ['指标', '数值'], rows: [['营收', '1000'], ['利润', '120']] },
    ],
  })
  return tools
}

describe('pptx_read structured query', () => {
  it('still returns plain text by default (backward compatible)', async () => {
    const file = deckPath('default.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_read', { file_path: file })
    expect(out).toContain('## Slide 1')
    expect(out).toContain('季度汇报')
    expect(out).not.toContain('**Summary:**')
  })

  it('reports per-slide summaries with include=summary', async () => {
    const file = deckPath('summary.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_read', { file_path: file, include: 'summary' })
    // pptxgenjs layout: title slide = title + subtitle (2 shapes); table slide = title + table
    expect(out).toMatch(/\*\*Summary:\*\* 2 text shape\(s\), 0 image\(s\), 0 table\(s\)/)
    expect(out).toMatch(/\*\*Summary:\*\* \d+ text shape\(s\), 0 image\(s\), 1 table\(s\)/)
  })

  it('reports shape layouts with positions in cm', async () => {
    const file = deckPath('layouts.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_read', { file_path: file, include: 'layouts' })
    expect(out).toContain('**Text shapes (name: pos — text):**')
    expect(out).toContain('cm')
    expect(out).toContain('季度汇报')
  })

  it('reports table dimensions and headers', async () => {
    const file = deckPath('tables.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_read', { file_path: file, include: 'tables' })
    expect(out).toContain('**Tables:**')
    expect(out).toMatch(/\d+ rows × 2 cols/)
    expect(out).toContain('指标 | 数值')
  })

  it('reports images when present', async () => {
    const file = deckPath('images.pptx')
    const tools = await makeDeck(file)
    // reuse the image-less deck is fine: images section is simply absent
    const out = await run(tools, 'pptx_read', { file_path: file, include: 'images' })
    // no images → no Images section (nothing to assert positively without an image asset)
    expect(out).not.toContain('**Images:**')
  })
})

describe('pptx_edit', () => {
  it('replaces text across all slides', async () => {
    const file = deckPath('edit-all.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_edit', {
      file_path: file,
      operations: [{ find: '营收', replace: 'Revenue' }],
    })
    expect(out).toContain('"营收" → "Revenue"')
    expect(out).toContain('replacement(s)')

    const read = await run(tools, 'pptx_read', { file_path: file })
    expect(read).toContain('Revenue增长 12%')
    expect(read).toContain('Revenue')
    expect(read).not.toContain('营收增长')
  })

  it('targets a single slide', async () => {
    const file = deckPath('edit-slide.pptx')
    const tools = await makeDeck(file)
    await run(tools, 'pptx_edit', {
      file_path: file,
      operations: [{ find: '季度汇报', replace: '年度汇报', slide: 1 }],
    })

    const read = await run(tools, 'pptx_read', { file_path: file })
    expect(read).toContain('年度汇报')
  })

  it('reports no match without touching the file', async () => {
    const file = deckPath('edit-nomatch.pptx')
    const tools = await makeDeck(file)
    const out = await run(tools, 'pptx_edit', {
      file_path: file,
      operations: [{ find: '不存在的文本', replace: 'x' }],
    })
    expect(out).toContain('No text matches found')

    const read = await run(tools, 'pptx_read', { file_path: file })
    expect(read).toContain('季度汇报')
  })

  it('writes to output_path when given', async () => {
    const file = deckPath('edit-out.pptx')
    const outFile = deckPath('edit-out-copy.pptx')
    const tools = await makeDeck(file)
    await run(tools, 'pptx_edit', {
      file_path: file,
      output_path: outFile,
      operations: [{ find: '成本下降', replace: '成本降低' }],
    })

    const orig = await run(tools, 'pptx_read', { file_path: file })
    const copy = await run(tools, 'pptx_read', { file_path: outFile })
    expect(orig).toContain('成本下降')
    expect(copy).toContain('成本降低')
  })

  it('rejects an out-of-range slide and a missing file', async () => {
    const file = deckPath('edit-err.pptx')
    const tools = await makeDeck(file)
    await expect(
      tools['pptx_edit']!.execute({ file_path: file, operations: [{ find: 'a', replace: 'b', slide: 99 }] }, noExec),
    ).rejects.toThrow(/slide 99 out of range/)
    await expect(
      tools['pptx_edit']!.execute({ file_path: join(dir, 'nope.pptx'), operations: [{ find: 'a', replace: 'b' }] }, noExec),
    ).rejects.toThrow('File not found')
  })
})
