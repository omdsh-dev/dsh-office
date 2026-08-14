// docx_create / docx_read tools for the DeepSeek Harness cordis tool model.
// Generation via the docx library (declarative document tree), extraction via
// mammoth. Mirrors the pdf_create content-block vocabulary.

import { existsSync } from 'node:fs'
import { writeFile } from 'node:fs/promises'
import { basename } from 'node:path'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx'
import mammoth from 'mammoth'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { textOutput } from './text-output.js'

interface DocxBlock {
  type?: string
  text?: string
  headers?: unknown[]
  rows?: unknown[][]
  items?: unknown[]
}

function toCellText(val: unknown): string {
  if (val === null || val === undefined) return ''
  return String(val)
}

function paragraphFor(block: DocxBlock): Paragraph | null {
  const text = block.text ?? ''
  switch (block.type) {
    case 'heading':
    case 'h1':
      return new Paragraph({ children: [new TextRun(text)], heading: HeadingLevel.HEADING_1 })
    case 'h2':
      return new Paragraph({ children: [new TextRun(text)], heading: HeadingLevel.HEADING_2 })
    case 'h3':
      return new Paragraph({ children: [new TextRun(text)], heading: HeadingLevel.HEADING_3 })
    case 'code':
      return new Paragraph({ children: [new TextRun({ text, font: 'Courier New' })] })
    default:
      return new Paragraph({ children: [new TextRun(text)] })
  }
}

/** Build one list block's bullet paragraphs (returned by reference in children). */
function listParagraphs(items: unknown[]): Paragraph[] {
  return items.map(item => new Paragraph({
    children: [new TextRun(toCellText(item))],
    bullet: { level: 0 },
  }))
}

function tableParagraphs(block: DocxBlock): Table | null {
  const headers = block.headers ?? []
  const rows = block.rows ?? []
  if (headers.length === 0 && rows.length === 0) return null
  const allRows: unknown[][] = headers.length > 0 ? [headers, ...rows] : rows
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: allRows.map(row => new TableRow({
      children: (Array.isArray(row) ? row : []).map(cell => new TableCell({
        children: [new Paragraph({ children: [new TextRun(toCellText(cell))] })],
      })),
    })),
  })
}

interface DocxCreateInput {
  destination_path: string
  title?: string
  content: unknown[]
}

async function docxCreate(params: DocxCreateInput): Promise<string> {
  const dest = params.destination_path
  if (!dest) throw new Error('destination_path is required')

  const children: (Paragraph | Table)[] = []
  if (params.title) {
    children.push(new Paragraph({
      children: [new TextRun({ text: params.title, bold: true, size: 36 })],
      alignment: 'center' as never,
    }))
    children.push(new Paragraph({ children: [new TextRun('')] }))
  }

  const blocks = Array.isArray(params.content) ? params.content : []
  for (const raw of blocks) {
    const block = (raw ?? {}) as DocxBlock
    if (block.type === 'table') {
      const table = tableParagraphs(block)
      if (table) children.push(table)
    } else if (block.type === 'list') {
      const items = Array.isArray(block.items) ? block.items : []
      children.push(...listParagraphs(items))
    } else {
      const p = paragraphFor(block)
      if (p) children.push(p)
    }
  }

  const doc = new Document({ sections: [{ children }] })
  const buf = await Packer.toBuffer(doc)
  await writeFile(dest, buf)
  return `📄 DOCX: Generated "${basename(dest)}" with ${children.length} block(s)\n   File: ${dest}`
}

const DOCX_READ_CHAR_CAP = 8000

async function docxRead(filePath: string): Promise<string> {
  if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
  const result = await mammoth.extractRawText({ path: filePath })
  const text = result.value ?? ''
  if (!text.trim()) return 'DOCX appears to contain no extractable text.'
  const truncated = text.length > DOCX_READ_CHAR_CAP
    ? text.slice(0, DOCX_READ_CHAR_CAP) + `\n\n... (truncated, ${text.length - DOCX_READ_CHAR_CAP} more chars.)`
    : text
  return truncated
}

export function registerDocxTools(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'docx_create',
    description: 'Generate a real .docx Word document from content blocks: {type:"heading"|"h2"|"h3"|"paragraph"|"table"|"code"|"list", text?, headers?, rows?, items?}. Optional title renders as a centered bold heading.',
    parameters: {
      destination_path: { type: 'string', required: true, description: 'Output .docx file path' },
      title: { type: 'string', description: 'Document title (centered, bold)' },
      content: {
        type: 'array', required: true,
        description: 'Content blocks: [{type, text?, headers?, rows?, items?}]',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            type: { type: 'string', enum: ['heading', 'h2', 'h3', 'paragraph', 'table', 'code', 'list'] as const },
            text: { type: 'string' },
            headers: { type: 'array', items: { type: 'string' } },
            rows: { type: 'array', items: { type: 'array' } },
            items: { type: 'array', items: { type: 'string' } },
          },
        },
      },
    },
    output: textOutput,
    execute: async (args) => ({ content: await docxCreate(args) }),
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'docx_read',
    description: 'Extract plain text from a .docx Word document for reading into context. Large documents are truncated at 8000 characters.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .docx file to read' },
    },
    output: textOutput,
    execute: async (args) => ({ content: await docxRead(args.file_path) }),
    isConcurrencySafe: () => true,
  }))
}
