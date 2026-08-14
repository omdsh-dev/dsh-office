// pptx round-trip tests: pptx_create → pptx_read, through the tool execute
// path. Covers title/content/table slides and speaker notes.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
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

describe('pptx tools', () => {
  it('pptx_create generates a deck and pptx_read extracts it', async () => {
    const { tools } = collect()
    const file = join(dir, 'deck.pptx')

    const created = await tools['pptx_create']!.execute({
      destination_path: file,
      title: 'Roadmap 2026',
      slides: [
        { type: 'title', title: 'Roadmap 2026', body: 'DeepSeek Harness ecosystem' },
        { type: 'content', title: 'Highlights', items: ['Plugin runtime', 'Office tools'] },
        { type: 'table', title: 'Milestones', headers: ['Quarter', 'Goal'], rows: [['Q1', 'Plugins'], ['Q2', 'Stable API']] },
      ],
    }, noExec) as { content: string }
    expect(created.content).toContain('3 slide(s)')
    expect(existsSync(file)).toBe(true)

    const read = await tools['pptx_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(read.content).toContain('Roadmap 2026')
    expect(read.content).toContain('## Slide 1')
    expect(read.content).toContain('## Slide 3')
    expect(read.content).toContain('Plugin runtime')
    expect(read.content).toContain('Q1')
  })

  it('pptx_create supports speaker notes and pptx_read includes them on demand', async () => {
    const { tools } = collect()
    const file = join(dir, 'notes.pptx')
    await tools['pptx_create']!.execute({
      destination_path: file,
      slides: [
        { type: 'content', title: 'Slide A', items: ['One'], notes: 'Remember to check the budget' },
      ],
    }, noExec)

    const plain = await tools['pptx_read']!.execute({ file_path: file }, noExec) as { content: string }
    expect(plain.content).not.toContain('budget')

    const withNotes = await tools['pptx_read']!.execute({ file_path: file, include_notes: true }, noExec) as { content: string }
    expect(withNotes.content).toContain('**Speaker notes:**')
    expect(withNotes.content).toContain('budget')
  })

  it('pptx_read on a missing file throws', async () => {
    const { tools } = collect()
    await expect(
      tools['pptx_read']!.execute({ file_path: join(dir, 'nope.pptx') }, noExec),
    ).rejects.toThrow(/file not found/)
  })
})
