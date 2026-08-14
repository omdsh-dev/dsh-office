// Real-harness integration test: assemble a minimal dsh cordis context with
// the published service packages (SystemPrompt → ToolRegistry → plugin) and
// drive tools through the real registry pipeline (schema validation, guards,
// execution), not through fakes.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Context } from '@deepseek-ai/cordis'
import ToolRegistry from '@deepseek-ai/dsh-tools'
import { SystemPrompt } from '@deepseek-ai/dsh-system-prompt'
import { name, inject, apply } from '../lib/index.js'

let ctx: Context
let dir: string
const signal = new AbortController().signal

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-integration-'))
  ctx = new Context()
  ctx.plugin(SystemPrompt)
  ctx.plugin(ToolRegistry, {})
  ctx.plugin({ name, inject, apply })
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

describe('real harness assembly', () => {
  it('registers all seven tools on the real registry', () => {
    const schemas = ctx.tools.schemas()
    const names = schemas.map(s => s.name).sort()
    expect(names).toEqual([
      'docx_create',
      'docx_read',
      'pdf_create',
      'pdf_read',
      'pptx_create',
      'pptx_read',
      'xlsx_edit',
      'xlsx_read',
      'xlsx_write',
    ])
    // Every schema carries parameters for the model.
    for (const s of schemas) {
      expect(s.parameters).toBeTruthy()
    }
  })

  it('runs xlsx_write through the registry and reads it back', async () => {
    const file = join(dir, 'book.xlsx')
    const write = await ctx.tools.execute({
      callId: 'it-xlsx-write',
      name: 'xlsx_write',
      signal,
      arguments: {
        file_path: file,
        data: [['Name', 'Score'], ['Alice', 92]],
        sheet_name: 'Results',
      },
    })
    expect(write.isError).toBe(false)
    if (!write.isError) {
      expect(write.value).toEqual({ content: expect.stringContaining('2 rows') })
    }
    expect(existsSync(file)).toBe(true)

    const read = await ctx.tools.execute({
      callId: 'it-xlsx-read',
      name: 'xlsx_read',
      signal,
      arguments: { file_path: file, sheet: 'Results' },
    })
    expect(read.isError).toBe(false)
    if (!read.isError) {
      expect(read.value).toEqual({ content: expect.stringContaining('Alice') })
    }
  })

  it('runs pdf_create through the registry and extracts it', async () => {
    const file = join(dir, 'doc.pdf')
    const created = await ctx.tools.execute({
      callId: 'it-pdf-create',
      name: 'pdf_create',
      signal,
      arguments: {
        destination_path: file,
        title: 'Integration Report',
        content: [{ type: 'paragraph', text: 'Assembled through the real registry.' }],
      },
    })
    expect(created.isError).toBe(false)
    expect(existsSync(file)).toBe(true)

    const read = await ctx.tools.execute({
      callId: 'it-pdf-read',
      name: 'pdf_read',
      signal,
      arguments: { file_path: file },
    })
    expect(read.isError).toBe(false)
    if (!read.isError) {
      expect(read.value).toEqual({ content: expect.stringContaining('Integration Report') })
    }
  })

  it('surfaces tool failures as registry failures, not exceptions', async () => {
    const result = await ctx.tools.execute({
      callId: 'it-xlsx-missing',
      name: 'xlsx_read',
      signal,
      arguments: { file_path: join(dir, 'nope.xlsx') },
    })
    expect(result.isError).toBe(true)
    if (result.isError) {
      expect(result.error).toBeTruthy()
    }
  })

  it('rejects unknown tool names', async () => {
    const result = await ctx.tools.execute({
      callId: 'it-unknown',
      name: 'no_such_tool',
      signal,
      arguments: {},
    })
    expect(result.isError).toBe(true)
  })
})
