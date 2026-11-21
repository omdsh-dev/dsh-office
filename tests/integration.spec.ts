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
import * as ToolExecutorNative from '@deepseek-ai/dsh-tool-executor-native'
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
  // 0.2.2-rc.5：`tools.execute` 要求装配期已装载一个 tool executor，否则抛
  // ToolExecutorUnavailableError（rc.2 下不需要）。补官方 native executor，
  // 让集成测试仍走真实的「schema 校验 → 执行 → 结果包装」链路而不是绕开。
  ctx.plugin(ToolExecutorNative, {})
  ctx.plugin({ name, inject, apply })
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

/**
 * 0.2.2-rc.5 起 `tools.execute` 是**批量**契约（`{ calls: [...] }` → `{ results }`），
 * 不再是单调用（`{ callId, name, arguments }` → 直接返回结果）。这里包一层，保持
 * 测试主体与断言不变：单调用当成只有一个元素的批，取首个结果。
 */
async function runTool(callId: string, name: string, args: Record<string, unknown>) {
  const batch = await ctx.tools.execute({
    calls: [{ id: callId, name, arguments: JSON.stringify(args) }],
    signal,
  })
  const first = batch.results[0]
  if (first === undefined) throw new Error(`no result for ${name}`)
  return first.result
}

describe('real harness assembly', () => {
  it('registers all twelve tools on the real registry', () => {
    const schemas = ctx.tools.schemas()
    const names = schemas.map(s => s.name).sort()
    expect(names).toEqual([
      'docx_create',
      'docx_read',
      'pdf_create',
      'pdf_merge',
      'pdf_read',
      'pdf_split',
      'pptx_create',
      'pptx_edit',
      'pptx_read',
      'xlsx_audit',
      'xlsx_edit',
      'xlsx_read',
      'xlsx_recalc',
      'xlsx_write',
    ])
    // Every schema carries parameters for the model.
    for (const s of schemas) {
      expect(s.parameters).toBeTruthy()
    }
  })

  it('runs xlsx_write through the registry and reads it back', async () => {
    const file = join(dir, 'book.xlsx')
    const write = await runTool('it-xlsx-write', 'xlsx_write', {
      file_path: file,
      data: [['Name', 'Score'], ['Alice', 92]],
      sheet_name: 'Results',
    })
    expect(write.isError).toBe(false)
    if (!write.isError) {
      expect(write.value).toEqual({ content: expect.stringContaining('2 rows') })
    }
    expect(existsSync(file)).toBe(true)

    const read = await runTool('it-xlsx-read', 'xlsx_read', { file_path: file, sheet: 'Results' })
    expect(read.isError).toBe(false)
    if (!read.isError) {
      expect(read.value).toEqual({ content: expect.stringContaining('Alice') })
    }
  })

  it('runs pdf_create through the registry and extracts it', async () => {
    const file = join(dir, 'doc.pdf')
    const created = await runTool('it-pdf-create', 'pdf_create', {
      destination_path: file,
      title: 'Integration Report',
      content: [{ type: 'paragraph', text: 'Assembled through the real registry.' }],
    })
    expect(created.isError).toBe(false)
    expect(existsSync(file)).toBe(true)

    const read = await runTool('it-pdf-read', 'pdf_read', { file_path: file })
    expect(read.isError).toBe(false)
    if (!read.isError) {
      expect(read.value).toEqual({ content: expect.stringContaining('Integration Report') })
    }
  })

  it('surfaces tool failures as registry failures, not exceptions', async () => {
    const result = await runTool('it-xlsx-missing', 'xlsx_read', { file_path: join(dir, 'nope.xlsx') })
    expect(result.isError).toBe(true)
    if (result.isError) {
      expect(result.error).toBeTruthy()
    }
  })

  it('rejects unknown tool names', async () => {
    const result = await runTool('it-unknown', 'no_such_tool', {})
    expect(result.isError).toBe(true)
  })
})
