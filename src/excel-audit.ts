// xlsx_audit / xlsx_recalc — formula quality gates for .xlsx workbooks.
//
// Heuristics adapted from the WPS Lingxi office-skills xlsx skill
// (recalc.py + audit.py, Apache-2.0 lineage via the Tianshu upstream):
//   • xlsx_recalc evaluates every formula with a lightweight pure-TS engine
//     and scans the results for error values (#REF!, #DIV/0!, #VALUE!, #N/A,
//     #NAME?, #NUM!) plus the "array formula trap" — constructs that Python
//     engines compute fine but Excel evaluates as #VALUE!.
//   • xlsx_audit statically inspects formula structure for problems a value
//     scan cannot see: aggregation ranges that miss an adjacent data row,
//     formulas overwritten by hardcoded values, inconsistent formulas within
//     a column, self-references, and literal division by zero.
//
// Both are zero-dependency on top of exceljs and intentionally heuristic —
// they flag suspicions for the model to confirm, never silently fix.

import type { Workbook, Worksheet } from 'exceljs'

// ── errors & values ────────────────────────────────────────────────

const EXCEL_ERRORS = ['#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NULL!', '#NUM!', '#N/A'] as const

type ErrorCode = (typeof EXCEL_ERRORS)[number]

interface Err {
  err: ErrorCode
}

type Value = number | string | boolean | null | Err

function isErr(v: Value): v is Err {
  return v !== null && typeof v === 'object' && 'err' in v
}

const ERR = {
  ref: { err: '#REF!' } as Err,
  div0: { err: '#DIV/0!' } as Err,
  value: { err: '#VALUE!' } as Err,
  name: { err: '#NAME?' } as Err,
  num: { err: '#NUM!' } as Err,
  na: { err: '#N/A' } as Err,
}

// ── reference parsing ──────────────────────────────────────────────

export interface Range {
  sheet: string // lowercased sheet name
  start: { col: number; row: number }
  end: { col: number; row: number }
}

const SHEET_NAME_RE = /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!/
const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/

/**
 * Parse an A1-style reference (optionally sheet-qualified, optionally a
 * range). Returns null when the text is not a plain reference.
 */
export function parseReference(text: string, currentSheet: string): Range | null {
  let rest = text.trim()
  let sheet = currentSheet
  const m = SHEET_NAME_RE.exec(rest)
  if (m) {
    sheet = (m[1] ?? m[2] ?? '').replace(/''/g, '')
    rest = rest.slice(m[0].length)
  }

  const parts = rest.split(':')
  if (parts.length > 2) return null
  const start = parseCellRef(parts[0]!)
  if (!start) return null

  let end = start
  if (parts.length === 2) {
    const endMatch = parseCellRef(parts[1]!)
    if (!endMatch) return null
    end = endMatch
  }

  return {
    sheet: sheet.toLowerCase(),
    start: { col: start.col, row: start.row },
    end: { col: Math.max(start.col, end.col), row: Math.max(start.row, end.row) },
  }
}

function parseCellRef(part: string): { col: number; row: number } | null {
  const m = CELL_RE.exec(part.trim())
  if (!m) return null
  const col = colToIndex(m[2] ?? '')
  const row = parseInt(m[4] ?? '0', 10)
  if (col <= 0 || row <= 0) return null
  return { col, row }
}

function colToIndex(col: string): number {
  let result = 0
  for (const ch of col.toUpperCase()) {
    result = result * 26 + (ch.charCodeAt(0) - 64)
  }
  return result
}

export function colToLetter(col: number): string {
  let n = col
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

// ── formula tokenizer / parser ─────────────────────────────────────

type TokenType = 'num' | 'str' | 'range' | 'op' | 'func' | 'lparen' | 'rparen' | 'comma' | 'percent' | 'eof'

interface Token {
  type: TokenType
  value: string | number
}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  const n = formula.length
  while (i < n) {
    const ch = formula[i]!
    if (ch === ' ' || ch === '\t') {
      i++
      continue
    }
    if (ch === '"') {
      let j = i + 1
      let s = ''
      while (j < n && formula[j] !== '"') {
        s += formula[j]
        j++
      }
      tokens.push({ type: 'str', value: s })
      i = j + 1
      continue
    }
    if (/[0-9.]/.test(ch)) {
      let j = i
      let sawE = false
      while (j < n) {
        const c = formula[j]!
        if (/[0-9.]/.test(c)) {
          j++
        } else if ((c === 'e' || c === 'E') && !sawE && /[0-9+-]/.test(formula[j + 1] ?? '')) {
          sawE = true
          j++
        } else if ((c === '+' || c === '-') && sawE && /[0-9]/.test(formula[j + 1] ?? '')) {
          j++
        } else {
          break
        }
      }
      let text = formula.slice(i, j)
      if (text.endsWith('.')) text = text.slice(0, -1)
      if (text.length > 0 && /[0-9]/.test(text)) {
        tokens.push({ type: 'num', value: Number(text) })
        i = j
        continue
      }
    }
    // quoted sheet reference: 'Sheet Name'!A1
    if (ch === "'") {
      let j = i + 1
      while (j < n && formula[j] !== "'") j++
      const sheetName = formula.slice(i + 1, j).replace(/''/g, '')
      j++
      if (formula[j] === '!') {
        j++
        const m = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/.exec(formula.slice(j))
        if (m) {
          const whole = m[0]!
          tokens.push({ type: 'range', value: `${sheetName}!${whole}` })
          i = j + whole.length
          continue
        }
      }
      tokens.push({ type: 'str', value: sheetName })
      i = j
      continue
    }
    // unquoted sheet-qualified reference: Sheet1!A1 / Sheet1!A1:B2
    const unquotedSheet = /^[A-Za-z_][A-Za-z0-9_.]*!/.exec(formula.slice(i))
    if (unquotedSheet) {
      const sheetPart = unquotedSheet[0]!
      const rest = formula.slice(i + sheetPart.length)
      const m = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/.exec(rest)
      if (m) {
        const whole = m[0]!
        const range = parseReference(`${sheetPart}${whole}`, '')
        if (range) {
          tokens.push({ type: 'range', value: `${sheetPart}${whole}` })
          i = i + sheetPart.length + whole.length
          continue
        }
      }
    }
    // identifier: function call (name followed by '(') or bare reference
    const ident = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(formula.slice(i))?.[0]
    if (ident) {
      const rest = formula.slice(i + ident.length)
      if (rest.startsWith('(') && /^[A-Za-z_]/.test(ident)) {
        tokens.push({ type: 'func', value: ident.toUpperCase() })
        i += ident.length
        continue
      }
      const rangeText = /^\$?[A-Za-z]{1,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d*)?/.exec(rest)?.[0] ?? ''
      const full = ident + rangeText
      if (parseReference(full, '')) {
        tokens.push({ type: 'range', value: full })
        i += full.length
        continue
      }
      // bare function without parentheses (TODAY, NOW, ...)
      tokens.push({ type: 'func', value: ident.toUpperCase() })
      i += ident.length
      continue
    }
    // operators
    const two = formula.slice(i, i + 2)
    if (two === '<>' || two === '<=' || two === '>=') {
      tokens.push({ type: 'op', value: two })
      i += 2
      continue
    }
    if ('+-*/^&()<>=,%'.includes(ch)) {
      const t: TokenType = ch === '(' ? 'lparen' : ch === ')' ? 'rparen' : ch === ',' ? 'comma' : ch === '%' ? 'percent' : 'op'
      tokens.push({ type: t, value: ch })
      i++
      continue
    }
    i++
  }
  tokens.push({ type: 'eof', value: '' })
  return tokens
}

// ── evaluator ──────────────────────────────────────────────────────

interface EvalContext {
  wb: Workbook
  formulaCache: Map<string, Value>
  evalStack: Set<string>
  steps: number
  maxSteps: number
  unknownFunctions: Set<string>
}

export interface RecalcResult {
  status: 'success' | 'errors_found'
  total_formulas: number
  total_errors: number
  error_summary: Record<string, { count: number; locations: string[] }>
  warnings: string[]
}

type RangeArg = { kind: 'range'; range: Range }
type Arg = Value | RangeArg

function isRangeArg(a: Arg): a is RangeArg {
  return a !== null && typeof a === 'object' && (a as RangeArg).kind === 'range'
}

type FnImpl = (args: Arg[], p: Parser) => Value

class Parser {
  private pos = 0
  readonly ctx: EvalContext

  constructor(
    private tokens: Token[],
    ctx: EvalContext,
    private sheet: string,
  ) {
    this.ctx = ctx
  }

  private peek(): Token {
    return this.tokens[this.pos]!
  }

  private next(): Token {
    return this.tokens[this.pos++]!
  }

  /** Parse one function argument: a bare range stays a RangeArg, anything
   *  else evaluates to a scalar. */
  private parseArg(): Arg {
    const t = this.peek()
    if (t.type === 'range') {
      this.next()
      const range = parseReference(String(t.value), this.sheet)
      return range ? { kind: 'range', range } : ERR.ref
    }
    return this.parseExpression()
  }

  parse(): Value {
    return this.parseExpression()
  }

  private parseExpression(): Value {
    return this.parseComparison()
  }

  private parseComparison(): Value {
    let left = this.parseConcat()
    for (;;) {
      const t = this.peek()
      if (t.type === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(String(t.value))) {
        this.next()
        const right = this.parseConcat()
        left = compareValues(left, right, String(t.value))
      } else {
        return left
      }
    }
  }

  private parseConcat(): Value {
    let left = this.parseAdditive()
    for (;;) {
      const t = this.peek()
      if (t.type === 'op' && t.value === '&') {
        this.next()
        const right = this.parseAdditive()
        left = concatValues(left, right)
      } else {
        return left
      }
    }
  }

  private parseAdditive(): Value {
    let left = this.parseMultiplicative()
    for (;;) {
      const t = this.peek()
      if (t.type === 'op' && (t.value === '+' || t.value === '-')) {
        this.next()
        const right = this.parseMultiplicative()
        left = t.value === '+'
          ? arith(left, right, (a, b) => a + b, ERR.value)
          : arith(left, right, (a, b) => a - b, ERR.value)
      } else {
        return left
      }
    }
  }

  private parseMultiplicative(): Value {
    let left = this.parsePower()
    for (;;) {
      const t = this.peek()
      if (t.type === 'op' && (t.value === '*' || t.value === '/')) {
        this.next()
        const right = this.parsePower()
        left = t.value === '*' ? arith(left, right, (a, b) => a * b, ERR.value) : divideValues(left, right)
      } else {
        return left
      }
    }
  }

  private parsePower(): Value {
    const left = this.parseUnary()
    const t = this.peek()
    if (t.type === 'op' && t.value === '^') {
      this.next()
      const right = this.parsePower()
      return arith(left, right, (a, b) => Math.pow(a, b), ERR.num)
    }
    return left
  }

  private parseUnary(): Value {
    const t = this.peek()
    if (t.type === 'op' && (t.value === '-' || t.value === '+')) {
      this.next()
      const v = this.parseUnary()
      if (isErr(v)) return v
      return typeof v === 'number' ? (t.value === '-' ? -v : v) : ERR.value
    }
    return this.parsePostfix()
  }

  private parsePostfix(): Value {
    let v = this.parseAtom()
    for (;;) {
      if (this.peek().type === 'percent') {
        this.next()
        if (isErr(v)) return v
        v = typeof v === 'number' ? v / 100 : ERR.value
      } else {
        return v
      }
    }
  }

  private parseAtom(): Value {
    const t = this.next()
    switch (t.type) {
      case 'num':
      case 'str':
        return t.value as number | string
      case 'range': {
        const range = parseReference(String(t.value), this.sheet)
        if (!range) return ERR.ref
        const ws = this.getSheet(range.sheet)
        if (!ws) return ERR.ref
        return this.cellValue(ws, range.start.col, range.start.row)
      }
      case 'func': {
        const name = String(t.value)
        const args: Arg[] = []
        if (this.peek().type === 'lparen') {
          this.next()
          if (this.peek().type !== 'rparen') {
            args.push(this.parseArg())
            while (this.peek().type === 'comma') {
              this.next()
              args.push(this.parseArg())
            }
          }
          this.next() // rparen
        }
        return this.callFunction(name, args)
      }
      case 'lparen': {
        const v = this.parseExpression()
        this.next() // rparen
        return v
      }
      default:
        return ERR.value
    }
  }

  getSheet(name: string): Worksheet | null {
    const ws = this.ctx.wb.getWorksheet(name)
    if (ws) return ws
    const capitalized = name[0]?.toUpperCase() + name.slice(1)
    return this.ctx.wb.getWorksheet(capitalized) ?? null
  }

  cellValue(ws: Worksheet, col: number, row: number): Value {
    const key = `${ws.name.toLowerCase()}!${col}:${row}`
    const cached = this.ctx.formulaCache.get(key)
    if (cached !== undefined) return cached
    if (this.ctx.evalStack.has(key)) return ERR.num // circular reference
    const raw = ws.getCell(row, col).value
    let v: Value
    if (raw === null || raw === undefined) {
      v = null
    } else if (typeof raw === 'object') {
      const obj = raw as { formula?: string; result?: unknown; text?: string }
      if (typeof obj.formula === 'string' && obj.formula.length > 0) {
        this.ctx.evalStack.add(key)
        const formulaText = obj.formula.startsWith('=') ? obj.formula.slice(1) : obj.formula
        v = this.evalFormulaText(formulaText, ws.name)
        this.ctx.evalStack.delete(key)
      } else if (obj.result !== undefined && obj.result !== null) {
        v = coerceScalar(obj.result)
      } else if (typeof obj.text === 'string') {
        v = obj.text
      } else {
        v = null
      }
    } else {
      v = coerceScalar(raw)
    }
    this.ctx.formulaCache.set(key, v)
    return v
  }

  private evalFormulaText(formula: string, sheetName: string): Value {
    this.ctx.steps++
    if (this.ctx.steps > this.ctx.maxSteps) return ERR.value
    const sub = new Parser(tokenize(formula), this.ctx, sheetName)
    return sub.parse()
  }

  private callFunction(name: string, args: Arg[]): Value {
    const fn = FUNCTIONS[name]
    if (!fn) {
      this.ctx.unknownFunctions.add(name)
      return ERR.name
    }
    try {
      return fn(args, this)
    } catch {
      return ERR.value
    }
  }

  /** Collect every cell value in a range (missing sheets yield no cells). */
  rangeValues(range: Range): Value[] {
    const ws = this.getSheet(range.sheet)
    if (!ws) return []
    const out: Value[] = []
    for (let c = range.start.col; c <= range.end.col; c++) {
      for (let r = range.start.row; r <= range.end.row; r++) {
        out.push(this.cellValue(ws, c, r))
      }
    }
    return out
  }
}

function collectNumeric(args: Arg[], p: Parser): { nums: number[]; err: Err | null } {
  const nums: number[] = []
  for (const a of args) {
    const v = scalarOf(a, p)
    if (isErr(v)) return { nums, err: v }
    if (isRangeArg(a)) {
      for (const rv of p.rangeValues(a.range)) {
        if (isErr(rv)) return { nums, err: rv }
        const n = asNumber(rv)
        if (n !== null) nums.push(n)
      }
    } else {
      const n = asNumber(v)
      if (n !== null) nums.push(n)
    }
  }
  return { nums, err: null }
}

/** Scalar value of a function argument: ranges collapse to their top-left cell. */
function scalarOf(a: Arg | undefined, p: Parser): Value {
  if (a === undefined || a === null) return null
  if (isRangeArg(a)) {
    const ws = p.getSheet(a.range.sheet)
    if (!ws) return ERR.ref
    return p.cellValue(ws, a.range.start.col, a.range.start.row)
  }
  return a
}

function asNumber(v: Value): number | null {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (typeof v === 'string') {
    const t = v.trim()
    if (t === '') return 0
    const n = Number(t)
    return Number.isFinite(n) ? n : null
  }
  return null
}

function asString(v: Value): string {
  if (v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return String(v)
}

function coerceScalar(raw: unknown): Value {
  if (typeof raw === 'number' || typeof raw === 'string' || typeof raw === 'boolean') return raw
  if (raw === null || raw === undefined) return null
  const s = String(raw)
  return EXCEL_ERRORS.includes(s as ErrorCode) ? ({ err: s as ErrorCode } as Err) : s
}

function arith(a: Value, b: Value, op: (x: number, y: number) => number, err: Err): Value {
  if (isErr(a)) return a
  if (isErr(b)) return b
  const an = a === null ? 0 : asNumber(a)
  const bn = b === null ? 0 : asNumber(b)
  if (an === null || bn === null) return err
  const r = op(an, bn)
  return Number.isFinite(r) ? r : err
}

function divideValues(a: Value, b: Value): Value {
  if (isErr(a)) return a
  if (isErr(b)) return b
  const an = a === null ? 0 : asNumber(a)
  const bn = b === null ? 0 : asNumber(b)
  if (an === null || bn === null) return ERR.value
  if (bn === 0) return ERR.div0
  const r = an / bn
  return Number.isFinite(r) ? r : ERR.div0
}

function compareValues(a: Value, b: Value, op: string): boolean {
  let cmp: number
  if (isErr(a) || isErr(b)) {
    cmp = isErr(a) && isErr(b) && a.err === b.err ? 0 : isErr(a) ? -1 : 1
  } else if (typeof a === 'number' && typeof b === 'number') {
    cmp = a < b ? -1 : a > b ? 1 : 0
  } else if (typeof a === 'string' && typeof b === 'string') {
    cmp = a < b ? -1 : a > b ? 1 : 0
  } else if (typeof a === 'boolean' && typeof b === 'boolean') {
    cmp = a === b ? 0 : a ? 1 : -1
  } else {
    const rank = (v: Value): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
    cmp = rank(a) - rank(b)
  }
  switch (op) {
    case '=': return cmp === 0
    case '<>': return cmp !== 0
    case '<': return cmp < 0
    case '>': return cmp > 0
    case '<=': return cmp <= 0
    case '>=': return cmp >= 0
  }
  return false
}

function concatValues(a: Value, b: Value): Value {
  if (isErr(a)) return a
  if (isErr(b)) return b
  return asString(a) + asString(b)
}

function truthy(v: Value): boolean {
  if (isErr(v)) return false
  if (v === null) return false
  if (typeof v === 'number') return v !== 0
  if (typeof v === 'boolean') return v
  return v !== ''
}

function criteriaMatches(value: Value, criteria: string): boolean {
  let c = criteria.trim()
  let op = '='
  const m = /^(<>|<=|>=|=|<|>)(.*)$/.exec(c)
  if (m) {
    op = m[1]!
    c = m[2]!.trim()
  }
  let target: Value
  if (c.startsWith('"') && c.endsWith('"')) target = c.slice(1, -1)
  else {
    const num = Number(c)
    target = Number.isFinite(num) && c.length > 0 ? num : c
  }
  if (value === null) return false
  return compareValues(value, target, op)
}

/** Sum over a range with a row offset for SUMIF/AVERAGEIF-style pairing. */
function pairedValue(range: Range, offset: number, p: Parser): Value | null {
  const ws = p.getSheet(range.sheet)
  if (!ws) return null
  const row = range.start.row + offset
  if (row > range.end.row) return null
  return p.cellValue(ws, range.start.col, row)
}

// ── function table ─────────────────────────────────────────────────

const FUNCTIONS: Record<string, FnImpl> = {
  SUM: (args, p) => {
    let sum = 0
    for (const a of args) {
      if (isErr(scalarOf(a, p))) return scalarOf(a, p)
      if (isRangeArg(a)) {
        for (const v of p.rangeValues(a.range)) {
          if (isErr(v)) return v
          const n = asNumber(v)
          if (n !== null) sum += n
        }
      } else {
        const n = asNumber(a)
        if (n !== null) sum += n
      }
    }
    return sum
  },
  AVERAGE: (args, p) => {
    const vals = collectNumeric(args, p)
    if (vals.err) return vals.err
    return vals.nums.length === 0 ? ERR.div0 : vals.nums.reduce((s, n) => s + n, 0) / vals.nums.length
  },
  COUNT: (args, p) => {
    let count = 0
    for (const a of args) {
      if (isErr(scalarOf(a, p))) return scalarOf(a, p)
      if (isRangeArg(a)) {
        for (const v of p.rangeValues(a.range)) {
          if (typeof v === 'number') count++
        }
      } else if (typeof a === 'number') {
        count++
      }
    }
    return count
  },
  COUNTA: (args, p) => {
    let count = 0
    for (const a of args) {
      if (isErr(scalarOf(a, p))) return scalarOf(a, p)
      if (isRangeArg(a)) {
        for (const v of p.rangeValues(a.range)) {
          if (v !== null) count++
        }
      } else if (a !== null) {
        count++
      }
    }
    return count
  },
  MIN: (args, p) => {
    const vals = collectNumeric(args, p)
    if (vals.err) return vals.err
    return vals.nums.length === 0 ? 0 : Math.min(...vals.nums)
  },
  MAX: (args, p) => {
    const vals = collectNumeric(args, p)
    if (vals.err) return vals.err
    return vals.nums.length === 0 ? 0 : Math.max(...vals.nums)
  },
  PRODUCT: (args, p) => {
    const vals = collectNumeric(args, p)
    if (vals.err) return vals.err
    return vals.nums.length === 0 ? 0 : vals.nums.reduce((s, n) => s * n, 1)
  },
  MEDIAN: (args, p) => {
    const vals = collectNumeric(args, p)
    if (vals.err) return vals.err
    if (vals.nums.length === 0) return ERR.num
    const sorted = [...vals.nums].sort((x, y) => x - y)
    const mid = Math.floor(sorted.length / 2)
    return sorted.length % 2 === 1 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2
  },
  ROUND: (args, p) => {
    const n = asNumber(scalarOf(args[0], p))
    const d = asNumber(scalarOf(args[1], p)) ?? 0
    if (n === null) return ERR.value
    const factor = Math.pow(10, Math.trunc(d))
    return Math.round(n * factor) / factor
  },
  ABS: (args, p) => {
    const n = asNumber(scalarOf(args[0], p))
    return n === null ? ERR.value : Math.abs(n)
  },
  MOD: (args, p) => {
    const a = asNumber(scalarOf(args[0], p))
    const b = asNumber(scalarOf(args[1], p))
    if (a === null || b === null) return ERR.value
    if (b === 0) return ERR.div0
    return a % b
  },
  INT: (args, p) => {
    const n = asNumber(scalarOf(args[0], p))
    return n === null ? ERR.value : Math.floor(n)
  },
  SQRT: (args, p) => {
    const n = asNumber(scalarOf(args[0], p))
    if (n === null) return ERR.value
    return n < 0 ? ERR.num : Math.sqrt(n)
  },
  POWER: (args, p) => {
    const a = asNumber(scalarOf(args[0], p))
    const b = asNumber(scalarOf(args[1], p))
    if (a === null || b === null) return ERR.value
    return Math.pow(a, b)
  },
  IF: (args, p) => truthy(scalarOf(args[0], p)) ? scalarOf(args[1], p) : scalarOf(args[2], p),
  AND: (args, p) => {
    for (const a of args) {
      if (isErr(scalarOf(a, p))) return scalarOf(a, p)
      if (!truthy(scalarOf(a, p))) return false
    }
    return true
  },
  OR: (args, p) => {
    for (const a of args) {
      if (isErr(scalarOf(a, p))) return scalarOf(a, p)
      if (truthy(scalarOf(a, p))) return true
    }
    return false
  },
  NOT: (args, p) => {
    const v = scalarOf(args[0], p)
    return isErr(v) ? v : !truthy(v)
  },
  IFERROR: (args, p) => {
    const v = scalarOf(args[0], p)
    return isErr(v) ? scalarOf(args[1], p) : v
  },
  IFNA: (args, p) => {
    const v = scalarOf(args[0], p)
    return isErr(v) && v.err === '#N/A' ? scalarOf(args[1], p) : v
  },
  LEFT: (args, p) => {
    const s = asString(scalarOf(args[0], p))
    const n = Math.max(0, Math.trunc(asNumber(scalarOf(args[1], p)) ?? 1))
    return s.slice(0, n)
  },
  RIGHT: (args, p) => {
    const s = asString(scalarOf(args[0], p))
    const n = Math.max(0, Math.trunc(asNumber(scalarOf(args[1], p)) ?? 1))
    return s.slice(-n)
  },
  MID: (args, p) => {
    const s = asString(scalarOf(args[0], p))
    const start = Math.max(0, Math.trunc(asNumber(scalarOf(args[1], p)) ?? 1) - 1)
    const n = Math.max(0, Math.trunc(asNumber(scalarOf(args[2], p)) ?? 0))
    return s.slice(start, start + n)
  },
  LEN: (args, p) => asString(scalarOf(args[0], p)).length,
  TRIM: (args, p) => asString(scalarOf(args[0], p)).replace(/\s+/g, ' ').trim(),
  UPPER: (args, p) => asString(scalarOf(args[0], p)).toUpperCase(),
  LOWER: (args, p) => asString(scalarOf(args[0], p)).toLowerCase(),
  CONCATENATE: (args, p) => args.map(a => asString(scalarOf(a, p))).join(''),
  CONCAT: (args, p) => args.map(a => asString(scalarOf(a, p))).join(''),
  VALUE: (args, p) => {
    const n = asNumber(scalarOf(args[0], p))
    return n === null ? ERR.value : n
  },
  VLOOKUP: (args, p) => {
    const lookup = scalarOf(args[0], p)
    const table = args[1]!
    if (!isRangeArg(table)) return ERR.value
    const colIdx = Math.trunc(asNumber(scalarOf(args[2], p)) ?? 1)
    const ws = p.getSheet(table.range.sheet)
    if (!ws) return ERR.na
    for (let r = table.range.start.row; r <= table.range.end.row; r++) {
      if (compareValues(p.cellValue(ws, table.range.start.col, r), lookup, '=')) {
        const targetCol = table.range.start.col + colIdx - 1
        if (targetCol > table.range.end.col || targetCol < table.range.start.col) return ERR.ref
        return p.cellValue(ws, targetCol, r)
      }
    }
    return ERR.na
  },
  INDEX: (args, p) => {
    const table = args[0]!
    if (!isRangeArg(table)) return ERR.value
    const ws = p.getSheet(table.range.sheet)
    if (!ws) return ERR.ref
    const rows = table.range.end.row - table.range.start.row + 1
    const cols = table.range.end.col - table.range.start.col + 1
    const rowIdx = Math.trunc(asNumber(scalarOf(args[1], p)) ?? 1)
    const colIdx = args.length > 2 ? Math.trunc(asNumber(scalarOf(args[2], p)) ?? 1) : 1
    if (rowIdx < 1 || rowIdx > rows || colIdx < 1 || colIdx > cols) return ERR.ref
    return p.cellValue(ws, table.range.start.col + colIdx - 1, table.range.start.row + rowIdx - 1)
  },
  MATCH: (args, p) => {
    const lookup = scalarOf(args[0], p)
    const table = args[1]!
    if (!isRangeArg(table)) return ERR.value
    const matchType = Math.trunc(asNumber(scalarOf(args[2], p)) ?? 1)
    if (matchType !== 0) return ERR.na // exact match only
    const ws = p.getSheet(table.range.sheet)
    if (!ws) return ERR.na
    for (let i = table.range.start.row; i <= table.range.end.row; i++) {
      if (compareValues(p.cellValue(ws, table.range.start.col, i), lookup, '=')) {
        return i - table.range.start.row + 1
      }
    }
    return ERR.na
  },
  SUMIF: (args, p) => {
    const range = args[0]!
    if (!isRangeArg(range)) return ERR.value
    const criteria = asString(scalarOf(args[1], p))
    const sumRange = isRangeArg(args[2]!) ? args[2].range : null
    let sum = 0
    for (let i = 0; i <= range.range.end.row - range.range.start.row; i++) {
      const test = pairedValue(range.range, i, p)
      if (test === null) break
      if (criteriaMatches(test, criteria)) {
        const target = sumRange ? pairedValue(sumRange, i, p) : test
        if (target === null) break
        const n = asNumber(target)
        if (n !== null) sum += n
      }
    }
    return sum
  },
  COUNTIF: (args, p) => {
    const range = args[0]!
    if (!isRangeArg(range)) return ERR.value
    const criteria = asString(scalarOf(args[1], p))
    let count = 0
    for (const v of p.rangeValues(range.range)) {
      if (criteriaMatches(v, criteria)) count++
    }
    return count
  },
  AVERAGEIF: (args, p) => {
    const range = args[0]!
    if (!isRangeArg(range)) return ERR.value
    const criteria = asString(scalarOf(args[1], p))
    const avgRange = isRangeArg(args[2]!) ? args[2].range : null
    let sum = 0
    let count = 0
    for (let i = 0; i <= range.range.end.row - range.range.start.row; i++) {
      const test = pairedValue(range.range, i, p)
      if (test === null) break
      if (criteriaMatches(test, criteria)) {
        const target = avgRange ? pairedValue(avgRange, i, p) : test
        if (target === null) break
        const n = asNumber(target)
        if (n !== null) {
          sum += n
          count++
        }
      }
    }
    return count === 0 ? ERR.div0 : sum / count
  },
  SUMIFS: (args, p) => {
    const sumRange = args[0]!
    if (!isRangeArg(sumRange)) return ERR.value
    let sum = 0
    for (let i = 0; i <= sumRange.range.end.row - sumRange.range.start.row; i++) {
      let match = true
      for (let ci = 1; ci + 1 < args.length; ci += 2) {
        const critRange = args[ci]!
        const crit = asString(scalarOf(args[ci + 1], p))
        if (!isRangeArg(critRange)) return ERR.value
        const test = pairedValue(critRange.range, i, p)
        if (test === null || !criteriaMatches(test, crit)) {
          match = false
          break
        }
      }
      if (match) {
        const v = pairedValue(sumRange.range, i, p)
        if (v === null) break
        const n = asNumber(v)
        if (n !== null) sum += n
      }
    }
    return sum
  },
  COUNTIFS: (args, p) => {
    const first = args[0]!
    if (!isRangeArg(first)) return ERR.value
    let count = 0
    for (let i = 0; i <= first.range.end.row - first.range.start.row; i++) {
      let match = true
      for (let ci = 0; ci + 1 < args.length; ci += 2) {
        const critRange = args[ci]!
        const crit = asString(scalarOf(args[ci + 1], p))
        if (!isRangeArg(critRange)) return ERR.value
        const test = pairedValue(critRange.range, i, p)
        if (test === null || !criteriaMatches(test, crit)) {
          match = false
          break
        }
      }
      if (match) count++
    }
    return count
  },
  AVERAGEIFS: (args, p) => {
    const avgRange = args[0]!
    if (!isRangeArg(avgRange)) return ERR.value
    let sum = 0
    let count = 0
    for (let i = 0; i <= avgRange.range.end.row - avgRange.range.start.row; i++) {
      let match = true
      for (let ci = 1; ci + 1 < args.length; ci += 2) {
        const critRange = args[ci]!
        const crit = asString(scalarOf(args[ci + 1], p))
        if (!isRangeArg(critRange)) return ERR.value
        const test = pairedValue(critRange.range, i, p)
        if (test === null || !criteriaMatches(test, crit)) {
          match = false
          break
        }
      }
      if (match) {
        const v = pairedValue(avgRange.range, i, p)
        if (v === null) break
        const n = asNumber(v)
        if (n !== null) {
          sum += n
          count++
        }
      }
    }
    return count === 0 ? ERR.div0 : sum / count
  },
  DATE: (args, p) => {
    const y = asNumber(scalarOf(args[0], p))
    const m = asNumber(scalarOf(args[1], p))
    const d = asNumber(scalarOf(args[2], p))
    if (y === null || m === null || d === null) return ERR.value
    return serialFromDate(new Date(Math.trunc(y), Math.trunc(m) - 1, Math.trunc(d)))
  },
  YEAR: (args, p) => {
    const s = serialToDate(scalarOf(args[0], p))
    return s === null ? ERR.value : s.getFullYear()
  },
  MONTH: (args, p) => {
    const s = serialToDate(scalarOf(args[0], p))
    return s === null ? ERR.value : s.getMonth() + 1
  },
  DAY: (args, p) => {
    const s = serialToDate(scalarOf(args[0], p))
    return s === null ? ERR.value : s.getDate()
  },
  TODAY: () => serialFromDate(new Date()),
  NOW: () => serialFromDate(new Date()),
}

const EPOCH = Date.UTC(1899, 11, 30)

function serialFromDate(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  return Math.round((utc - EPOCH) / 86400000)
}

function serialToDate(v: Value): Date | null {
  const n = asNumber(v)
  if (n === null) return null
  return new Date(EPOCH + Math.round(n) * 86400000)
}

// ── recalc: evaluate the whole workbook ────────────────────────────

export async function recalcWorkbook(wb: Workbook, timeoutSec = 30): Promise<RecalcResult> {
  const ctx: EvalContext = {
    wb,
    formulaCache: new Map(),
    evalStack: new Set(),
    steps: 0,
    maxSteps: Math.max(1000, Math.min(2_000_000, Math.floor(timeoutSec * 200_000))),
    unknownFunctions: new Set(),
  }

  const errorSummary: Record<string, { count: number; locations: string[] }> = {}
  let totalFormulas = 0

  const record = (err: ErrorCode, location: string): void => {
    const entry = (errorSummary[err] ??= { count: 0, locations: [] })
    entry.count++
    if (entry.locations.length < 20) entry.locations.push(location)
  }

  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const raw = cell.value
        if (raw === null || raw === undefined || typeof raw !== 'object') return
        const obj = raw as { formula?: string; result?: unknown }
        if (typeof obj.formula !== 'string' || obj.formula.length === 0) return
        totalFormulas++
        const key = `${ws.name.toLowerCase()}!${cell.col}:${cell.row}`
        ctx.evalStack.add(key)
        const formulaText = obj.formula.startsWith('=') ? obj.formula.slice(1) : obj.formula
        let v: Value
        try {
          v = new Parser(tokenize(formulaText), ctx, ws.name).parse()
        } catch {
          v = ERR.value
        }
        ctx.evalStack.delete(key)
        ctx.formulaCache.set(key, v)
        if (isErr(v)) {
          record(v.err, `${ws.name}!${cell.address}`)
        }
      })
    })
  }

  const warnings: string[] = []
  if (ctx.unknownFunctions.size > 0) {
    warnings.push(`Unsupported functions (evaluated as #NAME?): ${[...ctx.unknownFunctions].sort().join(', ')}`)
  }

  const totalErrors = Object.values(errorSummary).reduce((s, e) => s + e.count, 0)
  return {
    status: totalErrors > 0 ? 'errors_found' : 'success',
    total_formulas: totalFormulas,
    total_errors: totalErrors,
    error_summary: errorSummary,
    warnings,
  }
}

// ── static audit (heuristics from the Lingxi audit.py) ─────────────

export interface AuditWarning {
  type: string
  count: number
  samples: string[]
}

export interface AuditResult {
  status: 'clean' | 'warnings_found'
  total_formulas: number
  summary: Record<string, number>
  warnings: AuditWarning[]
}

const ARRAY_FORMULA_AGGREGATES =
  'SUM|AVERAGE|MEDIAN|MAX|MIN|COUNT|COUNTA|STDEV|STDEVP|VAR|VARP|PRODUCT|LARGE|SMALL'
const ARRAY_TRAP_RE = new RegExp(
  `\\b(?:${ARRAY_FORMULA_AGGREGATES})\\s*\\(\\s*(?:@?IF|IFERROR|AND|OR)\\s*\\(`,
  'i',
)
const AGGREGATE_RE = new RegExp(`\\b(?:${ARRAY_FORMULA_AGGREGATES}|IF|IFERROR)\\s*\\(`, 'gi')

/** Normalize a formula for structural comparison (case/whitespace/numbers/strings/refs). */
function normalizeFormula(text: string): string {
  return text
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/"[^"]*"/g, 'S')
    .replace(/\b\d+(\.\d+)?\b/g, 'N')
    .replace(/\$?[a-z]{1,3}\$?\d+/g, 'R')
}

function refsInFormula(text: string): Range[] {
  const refs: Range[] = []
  const re = new RegExp("(?:(?:'((?:[^']|'')+)'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\\$?[A-Za-z]{1,3}\\$?\\d+)(?::(\\$?[A-Za-z]{1,3}\\$?\\d+))?", 'g')
  let m: RegExpExecArray | null
  while ((m = re.exec(text)) !== null) {
    const start = m[2]!
    const end = m[3]
    const sheet = (m[1] ?? '').replace(/''/g, '')
    const r = parseReference(end ? `${start}:${end}` : start, sheet || '')
    if (r) refs.push(r)
  }
  return refs
}

export async function auditWorkbook(wb: Workbook): Promise<AuditResult> {
  const summary: Record<string, number> = {}
  const warningsMap = new Map<string, AuditWarning>()
  let totalFormulas = 0

  const addWarning = (type: string, location: string): void => {
    summary[type] = (summary[type] ?? 0) + 1
    const w = warningsMap.get(type) ?? { type, count: 0, samples: [] }
    w.count++
    if (w.samples.length < 5) w.samples.push(location)
    warningsMap.set(type, w)
  }

  const columnFormulas = new Map<string, Map<string, Array<{ row: number; text: string }>>>()
  const columnValues = new Map<string, Map<string, Map<number, number>>>()

  for (const ws of wb.worksheets) {
    const sheetKey = ws.name.toLowerCase()
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const raw = cell.value
        const colKey = `${sheetKey}|${colToLetter(Number(cell.col))}`
        if (raw === null || raw === undefined) return
        if (typeof raw === 'object') {
          const obj = raw as { formula?: string; result?: unknown }
          if (typeof obj.formula === 'string' && obj.formula.length > 0) {
            const text = obj.formula.startsWith('=') ? obj.formula.slice(1) : obj.formula
            totalFormulas++
            const bySheet = columnFormulas.get(colKey) ?? new Map()
            const list = bySheet.get(sheetKey) ?? []
            list.push({ row: cell.row, text })
            bySheet.set(sheetKey, list)
            columnFormulas.set(colKey, bySheet)

            if (typeof obj.result === 'string' && EXCEL_ERRORS.includes(obj.result as ErrorCode)) {
              addWarning('cached_error_value', `${ws.name}!${cell.address}`)
            }
            if (ARRAY_TRAP_RE.test(text)) {
              addWarning('array_formula_risk', `${ws.name}!${cell.address}`)
            }
            if (refsInFormula(text).some(
              r => (r.sheet === '' || r.sheet === sheetKey) && r.start.col === Number(cell.col) && r.start.row === Number(cell.row),
            )) {
              addWarning('self_reference', `${ws.name}!${cell.address}`)
            }
            if (/\/\s*0([^.\d]|$)/.test(text)) {
              addWarning('division_by_zero', `${ws.name}!${cell.address}`)
            }
            AGGREGATE_RE.lastIndex = 0
            if (AGGREGATE_RE.test(text)) {
              for (const ref of refsInFormula(text)) {
                if (ref.end.row > ref.start.row && ref.start.row > 1) {
                  const wsRef = ref.sheet === '' ? ws : (wb.getWorksheet(ref.sheet) ?? wb.getWorksheet(ref.sheet[0]!.toUpperCase() + ref.sheet.slice(1)))
                  if (!wsRef) {
                    addWarning('bad_reference', `${ws.name}!${cell.address}`)
                    continue
                  }
                  const above = wsRef.getCell(ref.start.row - 1, ref.start.col).value
                  const aboveRaw = above !== null && above !== undefined
                    ? (typeof above === 'object' && 'result' in above ? (above as { result?: unknown }).result : above)
                    : null
                  const isFormula = typeof above === 'object' && 'formula' in (above as object)
                  const isNumberAbove = typeof aboveRaw === 'number'
                  const isNumericString = typeof aboveRaw === 'string' && aboveRaw.trim() !== '' && Number.isFinite(Number(aboveRaw))
                  if (!isFormula && (isNumberAbove || isNumericString)) {
                    addWarning('range_gap', `${ws.name}!${cell.address} (missing row ${ref.start.row - 1}?)`)
                  }
                }
              }
            }
          } else if (typeof obj.result === 'number') {
            const bySheet = columnValues.get(colKey) ?? new Map()
            const rows = bySheet.get(sheetKey) ?? new Map()
            rows.set(cell.row, obj.result)
            bySheet.set(sheetKey, rows)
            columnValues.set(colKey, bySheet)
          }
        } else if (typeof raw === 'number') {
          const bySheet = columnValues.get(colKey) ?? new Map()
          const rows = bySheet.get(sheetKey) ?? new Map()
          rows.set(cell.row, raw)
          bySheet.set(sheetKey, rows)
          columnValues.set(colKey, bySheet)
        }
      })
    })
  }

  // overwritten formulas: bare numbers adjacent to a run of ≥2 formulas in a column
  for (const [colKey, bySheet] of columnFormulas) {
    const sep = colKey.indexOf('|')
    const colLetter = colKey.slice(sep + 1)
    for (const [sheetKey, list] of bySheet) {
      if (list.length < 2) continue
      const sorted = [...list].sort((a, b) => a.row - b.row)
      const valRows = columnValues.get(colKey)?.get(sheetKey)
      if (!valRows) continue
      const firstRow = sorted[0]!.row
      const lastRow = sorted[sorted.length - 1]!.row
      for (const probe of [firstRow - 1, lastRow + 1]) {
        if (valRows.has(probe)) {
          addWarning('possible_overwrite', `${sheetKey}!${colLetter}${probe}`)
        }
      }
    }
  }

  // inconsistent formulas within a column (minority structure differs)
  for (const [colKey, bySheet] of columnFormulas) {
    const sep = colKey.indexOf('|')
    const colLetter = colKey.slice(sep + 1)
    for (const [sheetKey, list] of bySheet) {
      if (list.length < 4) continue
      const buckets = new Map<string, number[]>()
      for (const item of list) {
        const norm = normalizeFormula(item.text)
        const arr = buckets.get(norm) ?? []
        arr.push(item.row)
        buckets.set(norm, arr)
      }
      if (buckets.size < 2) continue
      const entries = [...buckets.entries()].sort((a, b) => b[1].length - a[1].length)
      const majority = entries[0]!
      for (const [, rows] of entries.slice(1)) {
        if (rows.length < majority[1].length && rows.length <= 3) {
          for (const row of rows) {
            addWarning('inconsistent_formula', `${sheetKey}!${colLetter}${row}`)
          }
        }
      }
    }
  }

  return {
    status: Object.keys(summary).length > 0 ? 'warnings_found' : 'clean',
    total_formulas: totalFormulas,
    summary,
    warnings: [...warningsMap.values()],
  }
}
