// pdf_merge / pdf_split — page-level PDF operations built on pdf-lib
// (pure JS, no native deps). Mirrors the page-operation surface of the
// WPS Lingxi office-skills pdf skill (pdf_cli.py: merge / split).
// Encryption/decryption are intentionally absent: pdf-lib has not shipped
// an encrypt API, and requiring a system qpdf would break the
// zero-native-dep property of this plugin.

import { existsSync } from 'node:fs'
import { readFile, writeFile, mkdir } from 'node:fs/promises'
import { join, basename } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { textOutput } from './text-output.js'

/** Parse a page spec like "1,3,5-7" into 1-based indices within 1..total. */
function parsePages(spec: string, total: number): number[] {
  const out: number[] = []
  const seen = new Set<number>()
  for (const part of spec.split(',')) {
    const p = part.trim()
    if (!p) continue
    const range = /^(\d+)(?:-(\d+))?$/.exec(p)
    if (!range) throw new Error(`invalid page spec: "${part}" (expected e.g. "1,3,5-7")`)
    const start = parseInt(range[1]!, 10)
    const end = range[2] ? parseInt(range[2]!, 10) : start
    if (start < 1 || end < start || end > total) {
      throw new Error(`page out of range: "${part}" (file has ${total} pages)`)
    }
    for (let i = start; i <= end; i++) {
      if (!seen.has(i)) {
        seen.add(i)
        out.push(i)
      }
    }
  }
  if (out.length === 0) throw new Error('empty page spec')
  return out
}

async function loadDoc(filePath: string, ignoreEncryption = false): Promise<PDFDocument> {
  if (!existsSync(filePath)) throw new Error(`File not found: ${filePath}`)
  const bytes = await readFile(filePath)
  try {
    return await PDFDocument.load(bytes, ignoreEncryption ? { ignoreEncryption: true } : undefined)
  } catch (err) {
    throw new Error(`failed to load PDF: ${(err as Error).message}`)
  }
}

async function pdfMerge(params: {
  files: string[]
  output_path: string
}): Promise<string> {
  const files = params.files
  if (!Array.isArray(files) || files.length < 2) {
    throw new Error('pdf_merge requires at least 2 input files')
  }
  const out = await PDFDocument.create()
  let total = 0
  for (const f of files) {
    const src = await loadDoc(f)
    const pages = await out.copyPages(src, src.getPageIndices())
    for (const page of pages) out.addPage(page)
    total += pages.length
  }
  await writeFile(params.output_path, await out.save())
  return `📄 PDF: Merged ${files.length} file(s) → ${total} pages\n   File: ${params.output_path}`
}

async function pdfSplit(params: {
  file_path: string
  output_dir: string
  pages?: string // "1,3,5-7"; default: every page into its own file
}): Promise<string> {
  const src = await loadDoc(params.file_path)
  const total = src.getPageCount()
  const indices = params.pages ? parsePages(params.pages, total) : Array.from({ length: total }, (_, i) => i + 1)

  await mkdir(params.output_dir, { recursive: true })
  const stem = basename(params.file_path).replace(/\.pdf$/i, '')
  const written: string[] = []
  for (const pageNo of indices) {
    const out = await PDFDocument.create()
    const [page] = await out.copyPages(src, [pageNo - 1])
    if (!page) throw new Error(`failed to copy page ${pageNo}`)
    out.addPage(page)
    const dest = join(params.output_dir, `${stem}_p${String(pageNo).padStart(String(total).length, '0')}.pdf`)
    await writeFile(dest, await out.save())
    written.push(dest)
  }
  return [
    `📄 PDF: Split ${indices.length} page(s) from "${basename(params.file_path)}" (${total} pages)`,
    ...written.map(p => `   ${p}`),
  ].join('\n')
}

export function registerPdfOpsTools(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'pdf_merge',
    description: 'Merge two or more PDF files into one, in the given order.',
    parameters: {
      files: { type: 'array', required: true, items: { type: 'string' }, description: 'Input PDF paths (at least 2), merged in order' },
      output_path: { type: 'string', required: true, description: 'Output .pdf file path' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await pdfMerge(args) }),
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'pdf_split',
    description: 'Split a PDF into single-page files, or extract specific pages (page spec like "1,3,5-7"; default: every page).',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Input .pdf file path' },
      output_dir: { type: 'string', required: true, description: 'Directory to write page files into (created if missing)' },
      pages: { type: 'string', description: 'Page spec, e.g. "1,3,5-7" (1-based; default: all pages)' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await pdfSplit(args) }),
    isConcurrencySafe: () => true,
  }))
}
