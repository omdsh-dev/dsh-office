/**
 * @module @huiliyi37/dsh-office
 *
 * Office document tools for DeepSeek Harness: generate, read, and edit
 * spreadsheets (.xlsx), PDFs, presentations (.pptx), and Word documents
 * (.docx).
 *
 * Ported from the Tianshu terminal coding agent's office plugins
 * (Apache-2.0 licensed upstream, https://github.com/Tianshu-Tui).
 *
 * Registered tools (all enabled by default; disable per family via config):
 *   - xlsx_read / xlsx_write / xlsx_edit / xlsx_recalc / xlsx_audit  (exceljs)
 *   - pdf_create / pdf_read / pdf_merge / pdf_split                  (pdfkit + pdf-parse + pdf-lib)
 *   - pptx_create / pptx_read / pptx_edit                            (pptxgenjs + jszip)
 *   - docx_create / docx_read                                        (docx + mammoth)
 */

import type { Context } from '@deepseek-ai/cordis'
import z from '@deepseek-ai/schemastery'
import { registerExcelTools } from './excel.js'
import { registerPdfTools } from './pdf.js'
import { registerPptTools } from './ppt.js'
import { registerDocxTools } from './docx.js'

export const name = 'dsh-office'

/** Requires the tool registry seam. */
export const inject = ['tools'] as const

/** Per-family enable switches. Families omitted from `enable` stay enabled. */
export interface Config {
  enable?: {
    xlsx?: boolean
    pdf?: boolean
    ppt?: boolean
    docx?: boolean
  }
}

/** Schemastery configuration for loader validation. */
export const Config: z<Config> = z.object({
  enable: z.object({
    xlsx: z.boolean().required(false),
    pdf: z.boolean().required(false),
    ppt: z.boolean().required(false),
    docx: z.boolean().required(false),
  }).required(false),
})

/** Register office tools on the shared tool registry, honoring family switches. */
export function apply(ctx: Context, config: Config = {}): void {
  const enable = config.enable ?? {}
  if (enable.xlsx !== false) registerExcelTools(ctx)
  if (enable.pdf !== false) registerPdfTools(ctx)
  if (enable.ppt !== false) registerPptTools(ctx)
  if (enable.docx !== false) registerDocxTools(ctx)
}
