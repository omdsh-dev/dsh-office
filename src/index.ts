/**
 * @module @omdsh/dsh-office
 *
 * Office document tools for DeepSeek Harness: generate, read, and edit
 * spreadsheets (.xlsx), PDFs, and presentations (.pptx).
 *
 * Ported from the Tianshu terminal coding agent's office plugins
 * (Apache-2.0 licensed upstream, https://github.com/Tianshu-Tui/Tianshu-Tui).
 *
 * Registered tools:
 *   - xlsx_read / xlsx_write / xlsx_edit  (exceljs)
 *   - pdf_create / pdf_read               (pdfkit + pdf-parse, CJK-aware)
 *   - pptx_create / pptx_read             (pptxgenjs + jszip)
 *   - docx_create / docx_read             (docx + mammoth)
 */

import type { Context } from '@deepseek-ai/cordis'
import { registerExcelTools } from './excel.js'
import { registerPdfTools } from './pdf.js'
import { registerPptTools } from './ppt.js'
import { registerDocxTools } from './docx.js'

export const name = 'dsh-office'

/** Requires the tool registry seam. */
export const inject = ['tools'] as const

/** Register all office tools on the shared tool registry. */
export function apply(ctx: Context): void {
  registerExcelTools(ctx)
  registerPdfTools(ctx)
  registerPptTools(ctx)
  registerDocxTools(ctx)
}
