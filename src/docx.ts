// docx_create / docx_read tools for the DeepSeek Harness cordis tool model.
// Generation via the docx library (declarative document tree), extraction via
// mammoth. Mirrors the pdf_create content-block vocabulary.
//
// Extra typography beyond the docx library's surface (text watermark) is
// applied as an OOXML post-process patch on the packed buffer, following the
// same approach as the WPS Lingxi office-skills docx skill (docx_patches.js).

import { existsSync } from 'node:fs'
import { writeFile } from 'node:fs/promises'
import { basename } from 'node:path'
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, type IShadingAttributesProperties,
} from 'docx'
import JSZip from 'jszip'
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
  stripe?: boolean
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

const STRIPE_HEADER_FILL = '2F5496' // deep blue header row
const STRIPE_EVEN_FILL = 'F2F6FC'   // very light tint for alternating rows

function cellShading(fill: string): IShadingAttributesProperties {
  return { type: ShadingType.CLEAR, fill }
}

function tableParagraphs(block: DocxBlock): Table | null {
  const headers = block.headers ?? []
  const rows = block.rows ?? []
  if (headers.length === 0 && rows.length === 0) return null
  const allRows: unknown[][] = headers.length > 0 ? [headers, ...rows] : rows
  const stripe = block.stripe === true
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: allRows.map((row, ri) => new TableRow({
      children: (Array.isArray(row) ? row : []).map((cell, ci) => {
        void ci
        return new TableCell({
          children: [new Paragraph({ children: [new TextRun(toCellText(cell))] })],
          shading: stripe
            ? (ri === 0
              ? cellShading(STRIPE_HEADER_FILL)
              : (ri % 2 === 0 ? cellShading(STRIPE_EVEN_FILL) : undefined))
            : undefined,
          ...(stripe && ri === 0 ? { margins: { top: 60, bottom: 60, left: 80, right: 80 } } : {}),
        })
      }),
    })),
  })
}

interface WatermarkOptions {
  text: string
  color?: string // hex without '#', default '000000'
  opacity?: number // 0..1, default 0.3
}

interface DocxCreateInput {
  destination_path: string
  title?: string
  content: unknown[]
  background_color?: string // hex without '#'
  watermark?: WatermarkOptions
}

// ── OOXML watermark patch ───────────────────────────────────────────

const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const VML_NS = 'urn:schemas-microsoft-com:vml'
const HEADER_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
const HEADER_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

/**
 * Inject a diagonal text watermark (VML pict in a dedicated header) into a
 * packed .docx buffer. Pure zip/XML surgery — the docx library has no
 * watermark surface. Returns the patched buffer.
 */
async function applyWatermark(buffer: Buffer, opts: WatermarkOptions): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer)

  // pick a free header part name (Word starts numbering at header1)
  let headerName = 'word/header1.xml'
  let n = 1
  while (zip.file(headerName)) {
    n++
    headerName = `word/header${n}.xml`
  }

  const fill = (opts.color ?? '000000').replace(/^#/, '').toUpperCase()
  const opacity = Math.min(1, Math.max(0, opts.opacity ?? 0.3)).toFixed(2)
  const escaped = opts.text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${WML_NS}" xmlns:v="${VML_NS}">
  <w:p>
    <w:r><w:pict>
      <v:shape id="PowerPlusWaterMarkObject" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;rotation:315;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="${fill}" stroked="f">
        <v:fill opacity="${opacity}"/>
        <v:textpath style="font-family:&quot;Microsoft YaHei&quot;;font-size:1pt" string="${escaped}"/>
      </v:shape>
    </w:pict></w:r>
  </w:p>
</w:hdr>`

  // 1. document.xml — add a headerReference before the closing sectPr
  const docXml = await zip.file('word/document.xml')!.async('string')
  const headerRef = '<w:headerReference w:type="default" r:id="rIdWatermark"/>'
  if (!docXml.includes('</w:sectPr>')) {
    throw new Error('watermark patch failed: no sectPr found in document.xml')
  }
  const patchedDoc = docXml.replace('</w:sectPr>', `${headerRef}</w:sectPr>`)

  // 2. document.xml.rels — add the header relationship
  const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string')
  const rel = `<Relationship Id="rIdWatermark" Type="${HEADER_REL_TYPE}" Target="header${n}.xml"/>`
  const patchedRels = relsXml.replace('</Relationships>', `${rel}</Relationships>`)

  // 3. [Content_Types].xml — register the header part
  const typesXml = await zip.file('[Content_Types].xml')!.async('string')
  const override = `<Override PartName="/word/header${n}.xml" ContentType="${HEADER_CONTENT_TYPE}"/>`
  const patchedTypes = typesXml.replace('</Types>', `${override}</Types>`)

  zip.file('word/document.xml', patchedDoc)
  zip.file('word/_rels/document.xml.rels', patchedRels)
  zip.file('[Content_Types].xml', patchedTypes)
  zip.file(headerName, headerXml)

  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  })
}

// ── docx_create / docx_read ────────────────────────────────────────

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

  const background = params.background_color
    ? { color: (params.background_color as string).replace(/^#/, '').toUpperCase() }
    : undefined
  const doc = new Document({ background, sections: [{ children }] })
  let buf = await Packer.toBuffer(doc)

  let extras = ''
  if (params.watermark?.text) {
    buf = await applyWatermark(buf, {
      text: params.watermark.text,
      color: params.watermark.color,
      opacity: params.watermark.opacity,
    })
    extras += `, watermark "${params.watermark.text}"`
  }
  if (background) extras += `, background ${background.color}`

  await writeFile(dest, buf)
  return `📄 DOCX: Generated "${basename(dest)}" with ${children.length} block(s)${extras}\n   File: ${dest}`
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
    description: 'Generate a real .docx Word document from content blocks: {type:"heading"|"h2"|"h3"|"paragraph"|"table"|"code"|"list", text?, headers?, rows?, items?, stripe?}. Optional title renders as a centered bold heading. Optional background_color fills every page (e.g. a cover). Optional watermark overlays diagonal text (e.g. 机密 / DRAFT). Tables with stripe:true get a shaded header row and alternating row tints.',
    parameters: {
      destination_path: { type: 'string', required: true, description: 'Output .docx file path' },
      title: { type: 'string', description: 'Document title (centered, bold)' },
      background_color: { type: 'string', description: 'Page background color, hex without "#" (e.g. "1A1A2E" for a dark cover)' },
      watermark: {
        type: 'object',
        additionalProperties: false,
        description: 'Diagonal text watermark overlay',
        properties: {
          text: { type: 'string', required: true, description: 'Watermark text, e.g. "机密" or "DRAFT"' },
          color: { type: 'string', description: 'Watermark color, hex without "#" (default black)' },
          opacity: { type: 'number', description: 'Opacity 0..1 (default 0.3)' },
        },
      },
      content: {
        type: 'array', required: true,
        description: 'Content blocks: [{type, text?, headers?, rows?, items?, stripe?}]',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            type: { type: 'string', enum: ['heading', 'h2', 'h3', 'paragraph', 'table', 'code', 'list'] as const },
            text: { type: 'string' },
            headers: { type: 'array', items: { type: 'string' } },
            rows: { type: 'array', items: { type: 'array' } },
            items: { type: 'array', items: { type: 'string' } },
            stripe: { type: 'boolean', description: 'Table: shaded header row + alternating row tints' },
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
