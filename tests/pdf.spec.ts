// pdf round-trip tests: pdf_create → pdf_read, through the tool execute
// path. CJK assertions are conditional on a resolvable system font so the
// suite stays green on font-less CI runners.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import { registerPdfTools } from '../src/pdf.js'
import { resolveCjkFont } from '../src/fonts.js'

let dir: string
let hasCjkFont = false

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-pdf-'))
  hasCjkFont = (await resolveCjkFont()) !== null
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

function collect(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerPdfTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

const noExec = {} as never

describe('pdf tools', () => {
  it('pdf_create generates a readable PDF with headings, table and list', async () => {
    const { tools } = collect()
    const file = join(dir, 'doc.pdf')

    const created = await tools['pdf_create']!.execute({
      destination_path: file,
      title: 'Quarterly Report',
      content: [
        { type: 'heading', text: 'Summary' },
        { type: 'paragraph', text: 'Revenue grew across all regions.' },
        { type: 'table', headers: ['Region', 'Revenue'], rows: [['APAC', '120'], ['EMEA', '80']] },
        { type: 'list', items: ['Alpha', 'Beta'] },
      ],
      page_numbers: true,
    }, noExec) as { content: string }
    expect(created.content).toContain('Generated')
    expect(existsSync(file)).toBe(true)

    const read = await tools['pdf_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(read.content).toContain('Quarterly Report')
    expect(read.content).toContain('Summary')
    expect(read.content).toContain('Revenue grew')
    expect(read.content).toContain('APAC')
    expect(read.content).toContain('Alpha')
  })

  it('pdf_create renders CJK when a system font exists', async () => {
    const { tools } = collect()
    const file = join(dir, 'cjk.pdf')
    const created = await tools['pdf_create']!.execute({
      destination_path: file,
      title: '中文报告',
      content: [{ type: 'paragraph', text: '你好，世界。' }],
    }, noExec) as { content: string }

    if (!hasCjkFont) {
      // No CJK font on this machine: generation must still succeed, with a warning.
      expect(created.content).toMatch(/⚠️|Generated/)
      return
    }
    expect(created.content).toContain('Generated')
    const read = await tools['pdf_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(read.content).toContain('中文报告')
    expect(read.content).toContain('你好')
  })

  it('pdf_read on a missing file throws', async () => {
    const { tools } = collect()
    await expect(
      tools['pdf_read']!.execute({ file_path: join(dir, 'nope.pdf') }, noExec),
    ).rejects.toThrow(/file not found/)
  })
})
