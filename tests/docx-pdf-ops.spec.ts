// docx watermark/stripe/background + pdf_merge/pdf_split tests, through the
// tool execute path exactly as the harness would call them.

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { mkdtemp, rm, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ToolDefinition } from '@deepseek-ai/dsh-tools'
import JSZip from 'jszip'
import { PDFDocument } from 'pdf-lib'
import { registerDocxTools } from '../src/docx.js'
import { registerPdfTools } from '../src/pdf.js'

let dir: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'dsh-office-docxpdf-'))
})

afterAll(async () => {
  await rm(dir, { recursive: true, force: true })
})

function collectDocx(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerDocxTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

function collectPdf(): { tools: Record<string, ToolDefinition> } {
  const tools: Record<string, ToolDefinition> = {}
  registerPdfTools({ tools: { register: (d: ToolDefinition) => { tools[d.name] = d; return () => {} } } } as never)
  return { tools }
}

const noExec = {} as never

async function run(tools: Record<string, ToolDefinition>, name: string, params: Record<string, unknown>): Promise<string> {
  const res = await tools[name]!.execute(params, noExec) as { content: string }
  return res.content
}

describe('docx_create enhancements', () => {
  it('injects a watermark header via OOXML patch', async () => {
    const { tools } = collectDocx()
    const file = join(dir, 'wm.docx')
    const out = await run(tools, 'docx_create', {
      destination_path: file,
      title: '机密报告',
      watermark: { text: '机密', opacity: 0.4 },
      content: [{ type: 'paragraph', text: '正文' }],
    })
    expect(out).toContain('watermark "机密"')
    expect(existsSync(file)).toBe(true)

    // the zip must now contain the header part, a rel, a content-type override,
    // and a headerReference inside document.xml
    const zip = await JSZip.loadAsync(await readFile(file))
    const headerName = Object.keys(zip.files).find(n => /^word\/header\d+\.xml$/.test(n))
    expect(headerName).toBeTruthy()
    const headerXml = await zip.file(headerName!)!.async('string')
    expect(headerXml).toContain('机密')
    expect(headerXml).toContain('rotation:315')
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('rIdWatermark')
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(relsXml).toContain('rIdWatermark')
    const typesXml = await zip.file('[Content_Types].xml')!.async('string')
    expect(typesXml).toContain('wordprocessingml.header+xml')
  })

  it('applies background_color via the document background', async () => {
    const { tools } = collectDocx()
    const file = join(dir, 'bg.docx')
    const out = await run(tools, 'docx_create', {
      destination_path: file,
      background_color: '1A1A2E',
      content: [{ type: 'paragraph', text: '封面页' }],
    })
    expect(out).toContain('background 1A1A2E')

    const zip = await JSZip.loadAsync(await readFile(file))
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('w:background')
    expect(docXml).toContain('1A1A2E')
  })

  it('renders a striped table with shading', async () => {
    const { tools } = collectDocx()
    const file = join(dir, 'stripe.docx')
    await run(tools, 'docx_create', {
      destination_path: file,
      content: [
        {
          type: 'table',
          stripe: true,
          headers: ['Name', 'Score'],
          rows: [['Alice', 92], ['Bob', 88]],
        },
      ],
    })

    const zip = await JSZip.loadAsync(await readFile(file))
    const docXml = await zip.file('word/document.xml')!.async('string')
    // header row shading + alternating tint
    expect(docXml).toContain('2F5496')
    expect(docXml).toContain('F2F6FC')
  })

  it('docx_read still extracts text from a watermarked file', async () => {
    const { tools } = collectDocx()
    const file = join(dir, 'wm2.docx')
    await run(tools, 'docx_create', {
      destination_path: file,
      watermark: { text: 'DRAFT' },
      content: [{ type: 'paragraph', text: '内容可读' }],
    })
    const read = await run(tools, 'docx_read', { file_path: file })
    expect(read).toContain('内容可读')
  })
})

describe('pdf_merge / pdf_split', () => {
  it('merges two PDFs into one', async () => {
    const { tools } = collectPdf()
    const a = join(dir, 'a.pdf')
    const b = join(dir, 'b.pdf')
    await run(tools, 'pdf_create', { destination_path: a, title: 'A', content: [{ type: 'paragraph', text: 'one' }] })
    await run(tools, 'pdf_create', { destination_path: b, title: 'B', content: [{ type: 'paragraph', text: 'two' }] })

    const merged = join(dir, 'm.pdf')
    const out = await run(tools, 'pdf_merge', { files: [a, b], output_path: merged })
    expect(out).toContain('Merged 2 file(s) → 2 pages')
    expect(existsSync(merged)).toBe(true)

    const doc = await PDFDocument.load(await readFile(merged))
    expect(doc.getPageCount()).toBe(2)
  })

  it('splits a PDF into single-page files', async () => {
    const { tools } = collectPdf()
    const src = join(dir, 'multi.pdf')
    await run(tools, 'pdf_create', {
      destination_path: src,
      title: 'Multi',
      content: Array.from({ length: 80 }, (_, i) => ({ type: 'paragraph', text: `filler line ${i + 1} for page overflow` })),
    })

    const outDir = join(dir, 'split')
    const out = await run(tools, 'pdf_split', { file_path: src, output_dir: outDir })
    expect(out).toContain('Split 2 page(s)')
    expect(existsSync(join(outDir, 'multi_p1.pdf'))).toBe(true)
    expect(existsSync(join(outDir, 'multi_p2.pdf'))).toBe(true)

    const doc = await PDFDocument.load(await readFile(join(outDir, 'multi_p1.pdf')))
    expect(doc.getPageCount()).toBe(1)
  })

  it('extracts a page subset with a page spec', async () => {
    const { tools } = collectPdf()
    const src = join(dir, 'three.pdf')
    await run(tools, 'pdf_create', {
      destination_path: src,
      title: 'Three',
      content: Array.from({ length: 120 }, (_, i) => ({ type: 'paragraph', text: `filler line ${i + 1} for page overflow` })),
    })

    const outDir = join(dir, 'subset')
    const out = await run(tools, 'pdf_split', { file_path: src, output_dir: outDir, pages: '1,3' })
    expect(out).toContain('Split 2 page(s)')
    expect(existsSync(join(outDir, 'three_p1.pdf'))).toBe(true)
    expect(existsSync(join(outDir, 'three_p3.pdf'))).toBe(true)
    expect(existsSync(join(outDir, 'three_p2.pdf'))).toBe(false)
  })

  it('rejects an out-of-range page spec', async () => {
    const { tools } = collectPdf()
    const src = join(dir, 'one.pdf')
    await run(tools, 'pdf_create', { destination_path: src, title: 'One', content: [] })
    await expect(
      tools['pdf_split']!.execute({ file_path: src, output_dir: join(dir, 'bad'), pages: '5' }, noExec),
    ).rejects.toThrow(/page out of range/)
  })

  it('pdf_merge fails with a clear message for a missing file', async () => {
    const { tools } = collectPdf()
    const ok = join(dir, 'ok.pdf')
    await run(tools, 'pdf_create', { destination_path: ok, title: 'OK', content: [] })
    await expect(
      tools['pdf_merge']!.execute({ files: [ok, join(dir, 'nope.pdf')], output_path: join(dir, 'x.pdf') }, noExec),
    ).rejects.toThrow('File not found')
  })
})
