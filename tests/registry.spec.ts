// Plugin assembly test: applying the plugin registers all seven office tools
// on the tool registry, with the documented name/inject surface.

import { describe, it, expect } from 'vitest'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import { apply, name, inject } from '../src/index.js'

function makeCtx(): { ctx: unknown; definitions: ToolDefinition[] } {
  const definitions: ToolDefinition[] = []
  const ctx = {
    tools: {
      register: (d: ToolDefinition): (() => void) => {
        definitions.push(d)
        return () => {}
      },
    },
  }
  return { ctx, definitions }
}

describe('dsh-office plugin', () => {
  it('exposes the expected cordis plugin surface', () => {
    expect(name).toBe('dsh-office')
    expect(inject).toEqual(['tools'])
  })

  it('registers all seven office tools', () => {
    const { ctx, definitions } = makeCtx()
    apply(ctx as never)
    const names = definitions.map(d => d.name).sort()
    expect(names).toEqual([
      'pdf_create',
      'pdf_read',
      'pptx_create',
      'pptx_read',
      'xlsx_edit',
      'xlsx_read',
      'xlsx_write',
    ])
  })

  it('every tool carries a description, parameters, output schema and execute', () => {
    const { ctx, definitions } = makeCtx()
    apply(ctx as never)
    for (const d of definitions) {
      expect(typeof d.name).toBe('string')
      expect(typeof d.description).toBe('string')
      expect(d.description.length).toBeGreaterThan(0)
      expect(d.parameters).toBeTruthy()
      expect(d.output.schema).toBeTruthy()
      expect(typeof d.execute).toBe('function')
    }
  })
})
