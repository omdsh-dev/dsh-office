// docx round-trip tests: docx_create → docx_read, through the tool execute path.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import { registerDocxTools } from '../src/docx.js'

let dir: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-docx-'))
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

function collect(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerDocxTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

const noExec = {} as never

describe('docx tools', () => {
  it('docx_create generates a document and docx_read extracts it', async () => {
    const { tools } = collect()
    const file = join(dir, 'doc.docx')

    const created = await tools['docx_create']!.execute({
      destination_path: file,
      title: 'Quarterly Report',
      content: [
        { type: 'heading', text: 'Summary' },
        { type: 'paragraph', text: 'Revenue grew across all regions.' },
        { type: 'table', headers: ['Region', 'Revenue'], rows: [['APAC', '120'], ['EMEA', '80']] },
        { type: 'list', items: ['Alpha', 'Beta'] },
        { type: 'code', text: 'const x = 1' },
      ],
    }, noExec) as { content: string }
    expect(created.content).toContain('Generated')
    expect(existsSync(file)).toBe(true)

    const read = await tools['docx_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(read.content).toContain('Quarterly Report')
    expect(read.content).toContain('Summary')
    expect(read.content).toContain('Revenue grew')
    expect(read.content).toContain('APAC')
    expect(read.content).toContain('Alpha')
    expect(read.content).toContain('const x = 1')
  })

  it('docx_create supports h2/h3 blocks', async () => {
    const { tools } = collect()
    const file = join(dir, 'heads.docx')
    await tools['docx_create']!.execute({
      destination_path: file,
      content: [
        { type: 'h2', text: 'Sub Section' },
        { type: 'h3', text: 'Detail' },
      ],
    }, noExec)
    const read = await tools['docx_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(read.content).toContain('Sub Section')
    expect(read.content).toContain('Detail')
  })

  it('docx_read on a missing file throws', async () => {
    const { tools } = collect()
    await expect(
      tools['docx_read']!.execute({ file_path: join(dir, 'nope.docx') }, noExec),
    ).rejects.toThrow(/file not found/)
  })
})
