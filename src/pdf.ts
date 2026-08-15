// pdf_create / pdf_read tools, ported from the Tianshu office-pdf plugin
// (Apache-2.0 licensed upstream) to the DeepSeek Harness cordis tool model.

import { writeFileSync, renameSync, readFileSync, existsSync } from 'node:fs'
import { basename } from 'node:path'
import PDFDocument from 'pdfkit'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { containsCjk, resolveCjkFont } from './fonts.js'
import type { ResolvedCjkFont } from './fonts.js'
import { textOutput } from './text-output.js'
import { registerPdfOpsTools } from './pdf-ops.js'

// ── Helpers ──────────────────────────────────────────────────────

function artifactHint(filePath: string, summary: string): string {
  return [
    `📄 PDF: ${summary}`,
    `   File: ${filePath}`,
    `   Use read_file to inspect, or open_path to view.`,
  ].join('\n')
}

function toCellText(val: unknown): string {
  if (val === null || val === undefined) return ''
  return String(val)
}

interface PdfBlock {
  type?: string
  text?: string
  headers?: unknown[]
  rows?: unknown[][]
  items?: unknown[]
  ordered?: boolean
}

interface PdfInput {
  title?: string
  content?: string | PdfBlock[]
  pageNumbers?: boolean
}

function collectText(input: PdfInput): string {
  const parts: string[] = []
  if (input.title) parts.push(input.title)
  const blocks = Array.isArray(input.content) ? input.content : []
  for (const b of blocks) {
    if (!b) continue
    if (b.text) parts.push(b.text)
    if (Array.isArray(b.headers)) parts.push(b.headers.map(toCellText).join(' '))
    if (Array.isArray(b.rows)) for (const r of b.rows) parts.push((Array.isArray(r) ? r : []).map(toCellText).join(' '))
    if (Array.isArray(b.items)) parts.push(b.items.map(toCellText).join(' '))
  }
  if (typeof input.content === 'string') parts.push(input.content)
  return parts.join('\n')
}

async function generatePdf(filePath: string, input: PdfInput): Promise<string[]> {
  const warnings: string[] = []

  // CJK glyphs are absent from the built-in fonts — resolve a system font.
  const cjkNeeded = containsCjk(collectText(input))
  let cjkFont: ResolvedCjkFont | null = null
  if (cjkNeeded) {
    cjkFont = await resolveCjkFont()
    if (!cjkFont) {
      warnings.push('未找到 CJK 字体，中文可能无法渲染 (no CJK font found on this system; Chinese text may not render)')
    }
  }

  const doc = new PDFDocument({ size: 'A4', margin: 50, bufferPages: !!input.pageNumbers })
  const buffers: Buffer[] = []
  // pdfkit's typed `on` overloads are narrower than the runtime event set;
  // route events through the generic emitter surface.
  const emitter = doc as unknown as Pick<NodeJS.EventEmitter, 'on'>

  // Body/heading font setters — code blocks always switch back via applyBody.
  const applyBody = (): void => {
    if (cjkFont) {
      if (cjkFont.name) doc.font(cjkFont.path, cjkFont.name)
      else doc.font(cjkFont.path)
    } else {
      doc.font('Helvetica')
    }
  }
  const applyHeading = (): void => {
    if (cjkFont) {
      const name = cjkFont.headingName || cjkFont.name
      if (name) doc.font(cjkFont.path, name)
      else doc.font(cjkFont.path)
    } else {
      doc.font('Helvetica')
    }
  }

  return new Promise<string[]>((resolve, reject) => {
    emitter.on('data', (chunk: Buffer) => buffers.push(Buffer.from(chunk)))
    emitter.on('end', () => {
      // 原子替换：同目录临时文件 + rename（跨文件系统会 EXDEV，故不用 os.tmpdir）
      const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`
      writeFileSync(tmp, Buffer.concat(buffers))
      renameSync(tmp, filePath)
      resolve(warnings)
    })
    emitter.on('error', reject)
    const { title, content } = input

    applyBody()

    // Title
    if (title) {
      applyHeading()
      doc.fontSize(20).text(title, { align: 'center' })
      applyBody()
      doc.moveDown(1.5)
    }

    // Content blocks
    if (Array.isArray(content)) {
      for (const block of content) {
        if (!block) continue

        if (block.type === 'heading' || block.type === 'h1') {
          doc.moveDown(0.5)
          applyHeading()
          doc.fontSize(16).text(block.text || '', { continued: false })
          applyBody()
          doc.moveDown(0.5)
        } else if (block.type === 'h2') {
          doc.moveDown(0.3)
          applyHeading()
          doc.fontSize(14).text(block.text || '', { continued: false })
          applyBody()
          doc.moveDown(0.3)
        } else if (block.type === 'h3') {
          applyHeading()
          doc.fontSize(12).text(block.text || '', { continued: false })
          applyBody()
          doc.moveDown(0.2)
        } else if (block.type === 'paragraph' || block.type === 'text') {
          doc.fontSize(10).text(block.text || '', { align: 'justify' })
          doc.moveDown(0.5)
        } else if (block.type === 'table') {
          drawTable(doc, block, applyBody)
          doc.moveDown(0.5)
        } else if (block.type === 'list') {
          drawList(doc, block)
          doc.moveDown(0.5)
        } else if (block.type === 'code') {
          doc.font('Courier').fontSize(8).text(block.text || '')
          applyBody()
          doc.moveDown(0.3)
        } else {
          // fallback: plain text
          doc.fontSize(10).text(block.text || String(block))
          doc.moveDown(0.3)
        }
      }
    } else if (typeof content === 'string') {
      doc.fontSize(10).text(content, { align: 'justify' })
    }

    // Footer page numbers — second pass over buffered pages.
    if (input.pageNumbers) {
      const range = doc.bufferedPageRange()
      for (let i = 0; i < range.count; i++) {
        doc.switchToPage(range.start + i)
        applyBody()
        const label = cjkNeeded
          ? `第 ${i + 1} 页 / 共 ${range.count} 页`
          : `Page ${i + 1} of ${range.count}`
        doc.fontSize(8).text(label, doc.page.margins.left, doc.page.height - doc.page.margins.bottom + 15, {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          align: 'center',
          lineBreak: false,
        })
      }
    }

    doc.end()
  })
}

function drawList(doc: PDFKit.PDFDocument, block: PdfBlock): void {
  const items = Array.isArray(block.items) ? block.items : []
  if (items.length === 0) return
  const ordered = !!block.ordered
  const left = doc.page.margins.left
  const usable = doc.page.width - left - doc.page.margins.right

  doc.fontSize(10)
  items.forEach((item, idx) => {
    const bullet = ordered ? `${idx + 1}.` : '•'
    const y = doc.y
    // hanging indent: bullet in the gutter, text body indented
    doc.text(bullet, left + 4, y, { lineBreak: false })
    doc.text(toCellText(item), left + 20, y, { width: usable - 20 })
  })
}

function drawTable(doc: PDFKit.PDFDocument, block: PdfBlock, applyBody: () => void): void {
  const rows = block.rows || []
  const headers = block.headers || []
  if (rows.length === 0 && headers.length === 0) return

  if (applyBody) applyBody()
  const allRows: unknown[][] = headers.length > 0 ? [headers, ...rows] : rows
  const colCount = Math.max(...allRows.map(r => Array.isArray(r) ? r.length : 0), 1)
  const colWidth = (doc.page.width - doc.page.margins.left - doc.page.margins.right) / colCount
  const rowHeight = 18
  const fontSize = 9

  for (let ri = 0; ri < allRows.length; ri++) {
    const row = allRows[ri]
    const y = doc.y
    let maxH = rowHeight

    for (let ci = 0; ci < colCount; ci++) {
      const x = doc.page.margins.left + ci * colWidth
      const text = toCellText(Array.isArray(row) ? row[ci] : '')
      doc.fontSize(fontSize).text(text, x + 2, y + 2, {
        width: colWidth - 4,
        height: rowHeight - 4,
        ellipsis: true,
      })
      void maxH
    }

    // Draw cell borders
    doc.lineWidth(0.5)
    for (let ci = 0; ci <= colCount; ci++) {
      doc.moveTo(doc.page.margins.left + ci * colWidth, y)
        .lineTo(doc.page.margins.left + ci * colWidth, y + rowHeight)
        .stroke()
    }
    doc.moveTo(doc.page.margins.left, y + rowHeight)
      .lineTo(doc.page.margins.left + colCount * colWidth, y + rowHeight)
      .stroke()
    if (ri === 0) {
      doc.moveTo(doc.page.margins.left, y)
        .lineTo(doc.page.margins.left + colCount * colWidth, y)
        .stroke()
    }

    doc.y = y + rowHeight
    if (doc.y > doc.page.height - doc.page.margins.bottom - 40) {
      doc.addPage()
    }
  }
}

// ── pdf_read ────────────────────────────────────────────────────

interface PdfPageText {
  page: number
  text: string
}

const PDF_READ_CHAR_CAP = 8000

async function extractPdfPages(filePath: string): Promise<PdfPageText[]> {
  const { default: pdfParse } = await import('pdf-parse') as unknown as {
    default: (
      buffer: Buffer,
      options?: {
        pagerender?: (pageData: { getTextContent(): Promise<{ items: Array<{ str: string }> }> }) => Promise<string> | string
      },
    ) => Promise<{ numpages: number }>
  }
  const buffer = readFileSync(filePath)
  // pdf-parse's bundled pdf.js flakes with 'bad XRef entry' on the first
  // parse(s) after an idle period — retry with a short backoff.
  const pages: PdfPageText[] = []
  let lastErr: unknown
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      await pdfParse(buffer, {
        pagerender: async (pageData) => {
          const t = await pageData.getTextContent()
          const text = t.items.map(i => i.str).join(' ').trim()
          pages.push({ page: pages.length + 1, text })
          return text
        },
      })
      return pages
    } catch (err) {
      pages.length = 0
      lastErr = err
      if (attempt < 2) await new Promise(r => setTimeout(r, 75 * (attempt + 1)))
    }
  }
  throw lastErr
}

// ── registration ─────────────────────────────────────────────────

export function registerPdfTools(ctx: Context): void {
  registerPdfOpsTools(ctx)
  ctx.tools.register(defineTool({
    name: 'pdf_create',
    description: 'Generate a real PDF with text, headings, tables, and lists. CJK text is rendered via an auto-detected system font (warns if none found). Content is an array of blocks: {type:"heading"|"h2"|"h3"|"paragraph"|"table"|"code"|"list", text?, headers?, rows?, items?, ordered?}.',
    parameters: {
      destination_path: { type: 'string', required: true, description: 'Output .pdf file path' },
      title: { type: 'string', description: 'Document title' },
      page_numbers: { type: 'boolean', description: 'Add centered footer page numbers ("Page X of Y" / "第 X 页 / 共 Y 页")' },
      content: {
        type: 'json', required: true,
        description: 'Content: a string, or an array of blocks [{type, text?, headers?, rows?, items?, ordered?}]',
      },
    },
    output: textOutput,
    execute: async (args) => {
      const dest = args.destination_path
      if (!dest) throw new Error('destination_path is required')
      try {
        const warnings = await generatePdf(dest, {
          title: args.title,
          content: args.content as string | PdfBlock[] | undefined,
          pageNumbers: args.page_numbers === true,
        })
        const name = basename(dest)
        const warnText = warnings.length > 0 ? `\n⚠️ ${warnings.join('\n⚠️ ')}` : ''
        return { content: artifactHint(dest, `Generated "${name}"`) + warnText }
      } catch (err) {
        throw new Error(`PDF generation failed: ${(err as Error).message}`)
      }
    },
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'pdf_read',
    description: 'Extract text from a PDF file for reading into context. Each page is emitted under a "--- Page N ---" marker. Use start_page/end_page to read a specific range; large documents are truncated at 8000 characters with a continuation hint.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .pdf file to read' },
      start_page: { type: 'integer', description: 'First page to read (1-based, default 1)' },
      end_page: { type: 'integer', description: 'Last page to read (1-based, default last page)' },
    },
    output: textOutput,
    execute: async (args) => {
      const fp = args.file_path
      if (!fp) throw new Error('file_path is required')
      if (!existsSync(fp)) throw new Error(`file not found: ${fp}`)

      try {
        const pages = await extractPdfPages(fp)
        if (pages.length === 0) {
          return { content: 'PDF appears to contain no extractable text (scanned image?).' }
        }
        const start = Math.max(1, args.start_page ?? 1)
        const end = Math.min(pages.length, args.end_page ?? pages.length)
        if (start > end) {
          throw new Error(`invalid page range: start_page ${start} > end_page ${end} (total ${pages.length} pages)`)
        }

        // Page markers + per-page text, truncated at the page boundary.
        const blocks: string[] = []
        let chars = 0
        for (let i = start - 1; i < end; i++) {
          const p = pages[i]
          if (!p) continue
          const block = p.text.length > 0
            ? `--- Page ${p.page} ---\n${p.text}`
            : `--- Page ${p.page} ---\n(empty)`
          if (chars + block.length > PDF_READ_CHAR_CAP && blocks.length > 0) break
          blocks.push(block)
          chars += block.length + 1
        }

        const shownEnd = start + blocks.length - 1
        const hints: string[] = []
        if (shownEnd < end) hints.push(`Showing pages ${start}-${shownEnd} of ${pages.length}. Continue with start_page: ${shownEnd + 1}.`)
        else if (end < pages.length) hints.push(`Showing pages ${start}-${end} of ${pages.length}. Continue with start_page: ${end + 1}.`)
        if (hints.length > 0) blocks.push(hints.join('\n'))
        return { content: blocks.join('\n\n') }
      } catch (err) {
        throw new Error(`PDF read failed: ${(err as Error).message}`)
      }
    },
    isConcurrencySafe: () => true,
  }))
}
