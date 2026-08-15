// pptx_create / pptx_read tools, ported from the Tianshu office-ppt plugin
// (Apache-2.0 licensed upstream) to the DeepSeek Harness cordis tool model.

import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename } from 'node:path'
import type JSZip from 'jszip'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { textOutput } from './text-output.js'

// pptxgenjs's default export merges a class with a namespace; keep the
// instance surface we actually use as an explicit structural type.
interface PptxSlide {
  background: { color: string }
  addNotes(text: string): void
  addText(text: string | Array<{ text: string; options: Record<string, unknown> }>, opts: Record<string, unknown>): void
  addImage(opts: { path: string; x: number; y: number; w: number; h: number }): void
  addTable(rows: Array<Array<{ text: string; options: Record<string, unknown> }>>, opts: Record<string, unknown>): void
  addChart(kind: string, data: Array<{ name: string; labels: string[]; values: number[] }>, opts: Record<string, unknown>): void
}

interface PptxGenJS {
  addSlide(): PptxSlide
  layout: string
  author: string
  title: string
  writeFile(opts: { fileName: string }): Promise<void>
  ChartType: { bar: string; line: string; pie: string }
}

// Reject oversized decks before JSZip loads the whole archive into memory.
const MAX_PPTX_SIZE = 100 * 1024 * 1024 // 100MiB

// ── Helpers ──────────────────────────────────────────────────────

function artifactHint(filePath: string, summary: string): string {
  return [
    `📊 PPTX: ${summary}`,
    `   File: ${filePath}`,
    `   Use open_path to view in PowerPoint/Keynote.`,
  ].join('\n')
}

/** pptxgenjs wants bare hex (no leading '#'). */
function hex(value: string): string {
  return String(value).replace(/^#/, '').toUpperCase()
}

// Default theme preserves the original hardcoded appearance.
const DEFAULT_THEME = {
  titleColor: '1F2937',
  textColor: '374151',
  bgColor: 'FFFFFF',
  accentColor: '6B7280',
  fontFace: undefined as string | undefined,
}

interface ThemeInput {
  titleColor?: string
  textColor?: string
  bgColor?: string
  accentColor?: string
  fontFace?: string
}

interface ResolvedTheme {
  titleColor: string
  textColor: string
  bgColor: string
  accentColor: string
  fontFace: string | undefined
}

function resolveTheme(theme: ThemeInput | undefined): ResolvedTheme {
  const t = { ...DEFAULT_THEME, ...(theme || {}) }
  return {
    titleColor: hex(t.titleColor),
    textColor: hex(t.textColor),
    bgColor: hex(t.bgColor),
    accentColor: hex(t.accentColor),
    fontFace: t.fontFace ? String(t.fontFace) : undefined,
  }
}

// ── Text extraction helpers (pptx_read) ─────────────────────────

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&')
}

// Per-slide XML cap: regex backtracking on pathological (unclosed-tag)
// input scales ~k×n — a crafted deck could pin the event loop for minutes.
const MAX_SLIDE_XML_SIZE = 50 * 1024 * 1024 // 50MiB per slide XML

/** Reverse of unescapeXml — re-escape user text for XML round-trips. */
function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/** EMU (DrawingML internal unit) → centimetres, 1 decimal. */
function emuToCm(emu: string | undefined): number | null {
  if (!emu) return null
  const n = Number(emu)
  if (!Number.isFinite(n)) return null
  return Math.round((n / 360000) * 10) / 10
}

interface ShapeLayout {
  name: string
  text: string
  pos: string // "x,y w×h (cm)"
}

/** Extract text shapes (<p:sp>) with name, position/size (cm) and text. */
function extractShapeLayouts(xml: string): ShapeLayout[] {
  const out: ShapeLayout[] = []
  const spRe = /<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g
  let m: RegExpExecArray | null
  while ((m = spRe.exec(xml)) !== null) {
    const body = m[1] ?? ''
    const name = body.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1] ?? ''
    const off = body.match(/<a:off\b[^>]*\bx="(\d+)"\s*y="(\d+)"\/>/)
    const ext = body.match(/<a:ext\b[^>]*\bcx="(\d+)"\s*cy="(\d+)"\/>/)
    const text = extractParagraphs(body).join(' | ')
    const x = off ? emuToCm(off[1]) : null
    const y = off ? emuToCm(off[2]) : null
    const w = ext ? emuToCm(ext[1]) : null
    const h = ext ? emuToCm(ext[2]) : null
    const pos = [x, y, w, h].every(v => v !== null)
      ? `${x},${y} ${w}×${h}cm`
      : '(no geometry)'
    out.push({ name: name || '(unnamed)', text, pos })
  }
  return out
}

interface ImageRef {
  name: string
  embed: string
  pos: string
}

/** Extract pictures (<p:pic>) with name, embed rId and position. */
function extractImages(xml: string): ImageRef[] {
  const out: ImageRef[] = []
  const picRe = /<p:pic\b[^>]*>([\s\S]*?)<\/p:pic>/g
  let m: RegExpExecArray | null
  while ((m = picRe.exec(xml)) !== null) {
    const body = m[1] ?? ''
    const name = body.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1] ?? ''
    const embed = body.match(/<a:blip\b[^>]*\br:embed="([^"]*)"/)?.[1] ?? ''
    const off = body.match(/<a:off\b[^>]*\bx="(\d+)"\s*y="(\d+)"\/>/)
    const ext = body.match(/<a:ext\b[^>]*\bcx="(\d+)"\s*cy="(\d+)"\/>/)
    const x = off ? emuToCm(off[1]) : null
    const y = off ? emuToCm(off[2]) : null
    const w = ext ? emuToCm(ext[1]) : null
    const h = ext ? emuToCm(ext[2]) : null
    const pos = [x, y, w, h].every(v => v !== null)
      ? `${x},${y} ${w}×${h}cm`
      : '(no geometry)'
    out.push({ name: name || '(unnamed)', embed: embed || '(no embed)', pos })
  }
  return out
}

interface TableRef {
  pos: string
  rows: number
  cols: number
  header: string
}

/** Extract tables (<a:tbl>) with frame position, dimensions and header row text. */
function extractTables(xml: string): TableRef[] {
  const out: TableRef[] = []
  const tblRe = /<a:tbl\b[^>]*>([\s\S]*?)<\/a:tbl>/g
  let m: RegExpExecArray | null
  while ((m = tblRe.exec(xml)) !== null) {
    const body = m[1] ?? ''
    const grid = body.match(/<a:gridCol\b/g)?.length ?? 0
    const rows = body.match(/<a:tr\b/g)?.length ?? 0
    // first row's cells
    const firstTr = body.match(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/)
    const header = firstTr ? extractParagraphs(firstTr[1] ?? '').join(' | ') : ''
    // frame geometry lives in the enclosing <p:graphicFrame>; walk back from
    // the table start to the nearest frame open tag
    const before = xml.slice(0, m.index)
    const frameOpen = before.lastIndexOf('<p:graphicFrame')
    let pos = '(no frame)'
    if (frameOpen >= 0) {
      const frameHead = xml.slice(frameOpen, m.index).slice(0, 600)
      const off = frameHead.match(/<a:off\b[^>]*\bx="(\d+)"\s*y="(\d+)"\s*\/>/)
      const ext = frameHead.match(/<a:ext\b[^>]*\bcx="(\d+)"\s*cy="(\d+)"\s*\/>/)
      const x = off ? emuToCm(off[1]) : null
      const y = off ? emuToCm(off[2]) : null
      const w = ext ? emuToCm(ext[1]) : null
      const h = ext ? emuToCm(ext[2]) : null
      if (x !== null && y !== null && w !== null && h !== null) {
        pos = `${x},${y} ${w}×${h}cm`
      }
    }
    out.push({ pos, rows, cols: grid, header: header || '(empty)' })
  }
  return out
}

/** Join <a:t> runs within each <a:p> paragraph; drop empty paragraphs. */
function extractParagraphs(xml: string): string[] {
  if (xml.length > MAX_SLIDE_XML_SIZE) {
    return ['[slide content too large to parse]']
  }
  const paragraphs: string[] = []
  const pRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g
  let m: RegExpExecArray | null
  while ((m = pRe.exec(xml)) !== null) {
    const runs: string[] = []
    const tRe = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g
    let t: RegExpExecArray | null
    while ((t = tRe.exec(m[1] ?? '')) !== null) runs.push(unescapeXml(t[1] ?? ''))
    const text = runs.join('').trim()
    if (text) paragraphs.push(text)
  }
  return paragraphs
}

function slideNumber(name: string): number {
  return Number(name?.match(/slide(\d+)\.xml$/)?.[1] ?? NaN)
}

/** Map image rIds to media targets via the slide's rels part. */
async function resolveImageTargets(
  zip: JSZip,
  num: number,
): Promise<Map<string, string>> {
  const map = new Map<string, string>()
  const relsFile = zip.file(`ppt/slides/_rels/slide${num}.xml.rels`)
  if (!relsFile) return map
  const rels = await relsFile.async('string')
  const re = /<Relationship\s+Id="([^"]+)"\s+Type="[^"]*\/image"[^>]*Target="([^"]*)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(rels)) !== null) {
    map.set(m[1]!, m[2]!)
  }
  return map
}

/** Resolve the notes part for a slide via its .rels, falling back to same-numbered notesSlide. */
async function findNotesPart(
  zip: JSZip,
  num: number,
): Promise<string | null> {
  const relsFile = zip.file(`ppt/slides/_rels/slide${num}.xml.rels`)
  if (relsFile) {
    const rels = await relsFile.async('string')
    const m = rels.match(/Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/)
    if (m) return `ppt/notesSlides/${m[1]}`
  }
  const fallback = `ppt/notesSlides/notesSlide${num}.xml`
  return zip.file(fallback) ? fallback : null
}

// ── Slide builders ──────────────────────────────────────────────

interface SlideDef {
  type?: string
  title?: string
  body?: string
  items?: string[] | string
  image?: string
  headers?: string[]
  rows?: unknown[][]
  notes?: string
  chart?: string
  data?: Array<{ name?: string; labels?: unknown[]; values?: number[] }>
}

function addSlide(
  pptx: PptxGenJS,
  slideDef: SlideDef,
  theme: ResolvedTheme,
): void {
  const slide = pptx.addSlide()
  slide.background = { color: theme.bgColor }
  const { type, title, body, items, image } = slideDef
  const font = { fontFace: theme.fontFace }

  // Speaker notes (supported on every slide type)
  if (slideDef.notes) slide.addNotes(String(slideDef.notes))

  if (type === 'title') {
    // Title slide
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 1.5, w: '90%', h: 1.5,
        fontSize: 40, bold: true, align: 'center', color: theme.titleColor, ...font,
      })
    }
    if (body) {
      slide.addText(body, {
        x: 1, y: 3.2, w: '80%', h: 1,
        fontSize: 18, align: 'center', color: theme.accentColor, ...font,
      })
    }
  } else if (type === 'section') {
    // Section divider
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 2, w: '90%', h: 1.5,
        fontSize: 36, bold: true, align: 'center', color: theme.titleColor, ...font,
      })
    }
  } else if (type === 'content' || !type) {
    // Content slide: title + bullet points
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 0.4, w: '90%', h: 0.8,
        fontSize: 36, bold: true, color: theme.titleColor, ...font,
      })
    }
    if (body) {
      slide.addText(body, {
        x: 0.7, y: 1.5, w: '85%', h: 4,
        fontSize: 14, color: theme.textColor, bullet: !!items, ...font,
      })
    }
    if (Array.isArray(items) && items.length > 0) {
      const listItems = items.map(i => ({ text: String(i), options: { fontSize: 14, bullet: true, color: theme.textColor, ...font } }))
      slide.addText(listItems, {
        x: 0.7, y: 1.5, w: '85%', h: 4,
      })
    }
  } else if (type === 'two-column') {
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 0.4, w: '90%', h: 0.8,
        fontSize: 36, bold: true, color: theme.titleColor, ...font,
      })
    }
    // Left column
    slide.addText(body || '', {
      x: 0.5, y: 1.5, w: 4.2, h: 4,
      fontSize: 12, color: theme.textColor, ...font,
    })
    // Right column
    if (items) {
      slide.addText(Array.isArray(items) ? items.map(i => ({ text: String(i), options: { fontSize: 12, bullet: true, color: theme.textColor, ...font } })) : [{ text: String(items), options: { fontSize: 12, color: theme.textColor, ...font } }], {
        x: 5.2, y: 1.5, w: 4.2, h: 4,
      })
    }
  } else if (type === 'image') {
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 0.4, w: '90%', h: 0.8,
        fontSize: 36, bold: true, color: theme.titleColor, ...font,
      })
    }
    if (image) {
      slide.addImage({ path: image, x: 1, y: 1.5, w: 8, h: 4.5 })
    }
  } else if (type === 'table') {
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 0.4, w: '90%', h: 0.8,
        fontSize: 36, bold: true, color: theme.titleColor, ...font,
      })
    }
    if (slideDef.headers && slideDef.rows) {
      // addTable expects rows as cell arrays: [headerCells, rowCells, ...].
      const rows: Array<Array<{ text: string; options: { bold?: boolean; fill?: string; color: string; fontFace?: string } }>> = [slideDef.headers.map(h => ({ text: String(h), options: { bold: true, fill: 'E5E7EB', color: theme.titleColor, ...font } }))]
      for (const row of slideDef.rows) {
        rows.push(row.map(cell => ({ text: String(cell ?? ''), options: { color: theme.textColor, ...font } })))
      }
      slide.addTable(rows, {
        x: 0.5, y: 1.5, w: '90%',
        border: { type: 'solid', pt: 0.5, color: 'D1D5DB' },
        colW: Array(slideDef.headers.length).fill(9 / slideDef.headers.length),
      })
    }
  } else if (type === 'chart') {
    if (title) {
      slide.addText(title, {
        x: 0.5, y: 0.4, w: '90%', h: 0.8,
        fontSize: 36, bold: true, color: theme.titleColor, ...font,
      })
    }
    const chartType = { bar: pptx.ChartType.bar, line: pptx.ChartType.line, pie: pptx.ChartType.pie }[slideDef.chart || 'bar'] || pptx.ChartType.bar
    const data = (Array.isArray(slideDef.data) ? slideDef.data : []).map(s => ({
      name: String(s?.name ?? ''),
      labels: (Array.isArray(s?.labels) ? s.labels : []).map(String),
      values: (Array.isArray(s?.values) ? s.values : []).map(Number),
    })).filter(s => s.values.length > 0)
    if (data.length > 0) {
      slide.addChart(chartType, data, {
        x: 0.5, y: 1.5, w: 9, h: 5,
        showLegend: data.length > 1,
        chartColors: [theme.accentColor, theme.titleColor, theme.textColor],
      })
    }
  }
}

// ── pptx_edit: find/replace text via <a:t> surgery ────────────────

interface EditOperation {
  find: string
  replace?: string
  slide?: number // 1-based; omitted = all slides
}

interface PptxEditInput {
  file_path: string
  output_path?: string
  operations: EditOperation[]
}

async function pptxEdit(params: PptxEditInput): Promise<string> {
  const filePath = params.file_path
  if (!existsSync(filePath)) throw new Error(`File not found: ${filePath}`)
  const operations = params.operations
  if (!Array.isArray(operations) || operations.length === 0) {
    throw new Error('Missing or invalid operations: expected non-empty array')
  }
  const stat = statSync(filePath)
  if (stat.size > MAX_PPTX_SIZE) {
    throw new Error(`file too large (${(stat.size / 1024 / 1024).toFixed(1)}MB > 100MB limit)`)
  }

  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(readFileSync(filePath))

  const slideNumbers = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .map(slideNumber)
    .sort((a, b) => a - b)
  if (slideNumbers.length === 0) {
    throw new Error(`no slides found in ${filePath} (not a valid .pptx?)`)
  }

  const run = new RegExp('<a:t([^>]*)>([\\s\\S]*?)<\\/a:t>', 'g')
  const report: string[] = []
  let totalReplacements = 0

  for (const op of operations) {
    if (!op || typeof op.find !== 'string' || op.find.length === 0) {
      throw new Error('each operation needs a non-empty find string')
    }
    const find = op.find
    const replace = typeof op.replace === 'string' ? op.replace : ''
    const targets = op.slide
      ? [op.slide].filter(n => Number.isInteger(n) && n >= 1 && n <= slideNumbers[slideNumbers.length - 1]!)
      : slideNumbers
    if (targets.length === 0) {
      throw new Error(`slide ${String(op.slide)} out of range (deck has ${slideNumbers.length} slides)`)
    }

    const perSlide: string[] = []
    for (const num of targets) {
      const path = `ppt/slides/slide${num}.xml`
      const file = zip.file(path)
      if (!file) continue
      let xml = await file.async('string')
      let count = 0
      xml = xml.replace(run, (full, attrs: string, content: string) => {
        const text = unescapeXml(content)
        if (!text.includes(find)) return full
        count++
        return `<a:t${attrs}>${escapeXml(text.split(find).join(replace))}</a:t>`
      })
      if (count > 0) {
        zip.file(path, xml)
        perSlide.push(`${num}×${count}`)
        totalReplacements += count
      }
    }

    if (perSlide.length > 0) {
      report.push(`"${find}" → "${replace}": ${perSlide.join(', ')}`)
    }
  }

  if (totalReplacements === 0) {
    return `📊 PPTX: No text matches found in ${basename(filePath)} — nothing changed`
  }

  const outPath = params.output_path || filePath
  await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }).then(buf => {
    const fs = require('node:fs') as typeof import('node:fs')
    fs.writeFileSync(outPath, buf)
  })
  const name = basename(filePath)
  return `📊 PPTX: Edited "${name}" — ${totalReplacements} replacement(s)\n${report.map(r => `   ${r}`).join('\n')}\n   File: ${outPath}`
}

// ── registration ─────────────────────────────────────────────────

export function registerPptTools(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'pptx_create',
    description: 'Generate a real .pptx file from slide definitions. Slides: [{type:"title"|"section"|"content"|"two-column"|"image"|"table"|"chart", title?, body?, items?, image?, headers?, rows?, chart?, data?, notes?}]. Optional theme: {titleColor, textColor, bgColor, accentColor, fontFace}.',
    parameters: {
      destination_path: { type: 'string', required: true, description: 'Output .pptx file path' },
      title: { type: 'string', description: 'Presentation title (used on first slide if no explicit title slide)' },
      theme: {
        type: 'object',
        additionalProperties: false,
        description: 'Optional visual theme. Colors are hex with or without #. Defaults preserve the built-in look.',
        properties: {
          titleColor: { type: 'string', description: 'Title text color (default 1F2937)' },
          textColor: { type: 'string', description: 'Body text color (default 374151)' },
          bgColor: { type: 'string', description: 'Slide background color (default FFFFFF)' },
          accentColor: { type: 'string', description: 'Muted/subtitle + chart accent color (default 6B7280)' },
          fontFace: { type: 'string', description: 'Font family for all text (e.g. "Microsoft YaHei")' },
        },
      },
      slides: {
        type: 'array', required: true,
        description: 'Slide definitions array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            type: { type: 'string', enum: ['title', 'section', 'content', 'two-column', 'image', 'table', 'chart'] as const },
            title: { type: 'string' },
            body: { type: 'string' },
            items: { type: 'array', items: { type: 'string' } },
            image: { type: 'string', description: 'Path to image file' },
            headers: { type: 'array', items: { type: 'string' } },
            rows: { type: 'array', items: { type: 'array' } },
            notes: { type: 'string', description: 'Speaker notes for this slide' },
            chart: { type: 'string', enum: ['bar', 'line', 'pie'] as const, description: 'Chart kind (type=chart)' },
            data: {
              type: 'array',
              description: 'Chart series (type=chart): [{name, labels: [...], values: [...]}]',
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  name: { type: 'string' },
                  labels: { type: 'array' },
                  values: { type: 'array', items: { type: 'number' } },
                },
              },
            },
          },
        },
      },
    },
    output: textOutput,
    execute: async (args) => {
      const dest = args.destination_path
      if (!dest) throw new Error('destination_path is required')

      try {
        const PptxGenJS = (await import('pptxgenjs')).default as unknown as new () => PptxGenJS
        const pptx = new PptxGenJS()
        const theme = resolveTheme(args.theme)

        pptx.layout = 'LAYOUT_WIDE'
        pptx.author = 'dsh-office'
        pptx.title = args.title || 'Presentation'

        const slides = args.slides || []
        for (const slideDef of slides) {
          addSlide(pptx, slideDef, theme)
        }

        await pptx.writeFile({ fileName: dest })
        const name = basename(dest)
        return { content: artifactHint(dest, `Generated "${name}" with ${slides.length} slide(s)`) }
      } catch (err) {
        throw new Error(`PPTX generation failed: ${(err as Error).message}`)
      }
    },
    isConcurrencySafe: () => true,
  }))
  ctx.tools.register(defineTool({
    name: 'pptx_edit',
    description: 'Edit text inside an existing .pptx by find/replace on text runs (<a:t> nodes), preserving all other layout and styling. Each operation: {find, replace, slide?} — slide is 1-based, omit to edit all slides. Saves back to file_path unless output_path is given. Perfect for fixing typos or updating numbers in a generated deck.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the existing .pptx file' },
      output_path: { type: 'string', description: 'Save to a different path instead of overwriting' },
      operations: {
        type: 'array', required: true,
        description: 'Ordered find/replace operations',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            find: { type: 'string', required: true, description: 'Text to find (plain text, not regex)' },
            replace: { type: 'string', description: 'Replacement text (empty string deletes the match)' },
            slide: { type: 'integer', description: '1-based slide number; omit to edit all slides' },
          },
        },
      },
    },
    output: textOutput,
    execute: async (args) => ({ content: await pptxEdit(args) }),
    isConcurrencySafe: () => false,
  }))
  ctx.tools.register(defineTool({
    name: 'pptx_read',
    description: 'Extract text from a .pptx file as markdown (## Slide N + paragraphs). Optionally includes speaker notes. Pass include to also get structure: "summary,layouts,images,tables" (e.g. shape names, positions in cm, image targets, table dimensions). Use it to review generated decks or analyze template layouts.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .pptx file' },
      include_notes: { type: 'boolean', description: 'Also extract speaker notes (default false)' },
      include: { type: 'string', description: 'Comma-separated structure to include per slide: summary, layouts, images, tables (e.g. "summary,images")' },
    },
    output: textOutput,
    execute: async (args) => {
      const filePath = args.file_path
      if (!filePath) throw new Error('file_path is required')
      if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)

      try {
        const stat = statSync(filePath)
        if (stat.size > MAX_PPTX_SIZE) {
          throw new Error(`file too large (${(stat.size / 1024 / 1024).toFixed(1)}MB > 100MB limit)`)
        }
        const JSZip = (await import('jszip')).default
        const zip = await JSZip.loadAsync(readFileSync(filePath))

        const slideNames = Object.keys(zip.files)
          .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
          .sort((a, b) => slideNumber(a) - slideNumber(b))
        if (slideNames.length === 0) {
          throw new Error(`no slides found in ${filePath} (not a valid .pptx?)`)
        }

        const includeNotes = args.include_notes === true
        const include = new Set((typeof args.include === 'string' ? args.include : '').split(',').map(s => s.trim()).filter(Boolean))
        const out: string[] = [`# ${basename(filePath)}`, '']
        for (let i = 0; i < slideNames.length; i++) {
          const num = slideNumber(slideNames[i] ?? '')
          const xml = await zip.file(slideNames[i] ?? '')?.async('string') ?? ''
          const paragraphs = extractParagraphs(xml)
          out.push(`## Slide ${i + 1}`)
          out.push(paragraphs.length > 0 ? paragraphs.join('\n\n') : '(no text)')

          if (include.size > 0) {
            if (include.has('layouts')) {
              const shapes = extractShapeLayouts(xml)
              if (shapes.length > 0) {
                out.push('', '**Text shapes (name: pos — text):**')
                for (const s of shapes) {
                  out.push(`- ${s.name}: ${s.pos} — ${s.text || '(no text)'}`)
                }
              }
            }
            if (include.has('images')) {
              const images = extractImages(xml)
              if (images.length > 0) {
                const targets = await resolveImageTargets(zip, num)
                out.push('', '**Images:**')
                for (const img of images) {
                  const target = targets.get(img.embed)
                  out.push(`- ${img.name}: ${img.pos}${target ? ` → ${target}` : ''}`)
                }
              }
            }
            if (include.has('tables')) {
              const tables = extractTables(xml)
              if (tables.length > 0) {
                out.push('', '**Tables:**')
                for (const t of tables) {
                  out.push(`- ${t.rows} rows × ${t.cols} cols at ${t.pos}; header: ${t.header}`)
                }
              }
            }
            if (include.has('summary')) {
              const shapes = extractShapeLayouts(xml)
              const images = extractImages(xml)
              const tables = extractTables(xml)
              out.push('', `**Summary:** ${shapes.length} text shape(s), ${images.length} image(s), ${tables.length} table(s)`)
            }
          }

          if (includeNotes) {
            const notesPart = await findNotesPart(zip, num)
            if (notesPart) {
              const notes = extractParagraphs(await zip.file(notesPart)?.async('string') ?? '')
              if (notes.length > 0) out.push('', '**Speaker notes:**', notes.join('\n\n'))
            }
          }
          out.push('')
        }

        return { content: out.join('\n') }
      } catch (err) {
        throw new Error(`PPTX read failed: ${(err as Error).message}`)
      }
    },
    isConcurrencySafe: () => true,
  }))
}
