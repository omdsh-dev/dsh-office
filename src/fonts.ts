// CJK font resolution for pdfkit: pdfkit's built-in Helvetica/Courier have no
// CJK glyphs, so when content contains CJK characters we locate a usable
// system font and register it explicitly (TTC files need a postscript name).
// Ported from the Tianshu office-pdf plugin (Apache-2.0 licensed upstream).

import { existsSync } from 'node:fs'

const CJK_RE = /[一-鿿㐀-䶿豈-﫿]/

export function containsCjk(text: string | undefined): boolean {
  return typeof text === 'string' && CJK_RE.test(text)
}

interface FontCandidate {
  path: string
  names: string[]
  headingNames?: string[]
}

// Ordered per-platform candidates. `names` are preferred postscript names
// inside a TTC collection; plain .ttf/.otf entries use an empty list (pdfkit
// loads them by path alone). `headingNames` optionally picks a heavier cut
// for headings, falling back to the body name.
const CJK_FONT_CANDIDATES: FontCandidate[] = [
  // macOS
  { path: '/System/Library/Fonts/PingFang.ttc', names: ['PingFangSC-Regular'] },
  { path: '/System/Library/Fonts/Hiragino Sans GB.ttc', names: ['HiraginoSansGB-W3'], headingNames: ['HiraginoSansGB-W6'] },
  { path: '/System/Library/Fonts/STHeiti Light.ttc', names: ['STHeitiSC-Light'] },
  // Windows
  { path: 'C:\\Windows\\Fonts\\msyh.ttc', names: ['MicrosoftYaHei'] },
  { path: 'C:\\Windows\\Fonts\\simhei.ttf', names: [] },
  { path: 'C:\\Windows\\Fonts\\simsun.ttc', names: ['SimSun'] },
  // Linux
  { path: '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', names: ['NotoSansCJKsc-Regular', 'NotoSansCJK-Regular'] },
  { path: '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf', names: [] },
  { path: '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', names: ['WenQuanYiZenHei'] },
]

function pickName(fonts: OpenedFont[], preferred: string[] | undefined): string | null {
  for (const n of preferred || []) {
    if (fonts.some(f => f.postscriptName === n)) return n
  }
  return null
}

export interface ResolvedCjkFont {
  path: string
  name: string | null
  headingName: string | null
}

interface OpenedFont {
  postscriptName?: string
}

interface FontkitLike {
  openSync(path: string): { fonts?: OpenedFont[] } & OpenedFont
}

/**
 * Resolve the first usable CJK font.
 * Returns { path, name, headingName } (name/headingName null for plain
 * ttf/otf), or null when no candidate exists / is loadable.
 */
export async function resolveCjkFont(): Promise<ResolvedCjkFont | null> {
  let fontkit: FontkitLike | null = null
  try {
    // CJS interop: the dynamic-import namespace carries the module as `.default`.
    const mod = (await import('fontkit')) as unknown as { default?: FontkitLike }
    fontkit = mod.default ?? null
  } catch {
    // fontkit ships with pdfkit; without it we can only use plain ttf/otf
  }

  for (const cand of CJK_FONT_CANDIDATES) {
    if (!existsSync(cand.path)) continue
    if (cand.names.length === 0) {
      return { path: cand.path, name: null, headingName: null }
    }
    if (!fontkit) continue
    try {
      const opened = fontkit.openSync(cand.path)
      const fonts: OpenedFont[] = opened.fonts && opened.fonts.length > 0 ? opened.fonts : [opened]
      const name = pickName(fonts, cand.names) || fonts[0]?.postscriptName || null
      if (!name) continue
      const headingName = pickName(fonts, cand.headingNames) || name
      return { path: cand.path, name, headingName }
    } catch {
      continue
    }
  }
  return null
}
