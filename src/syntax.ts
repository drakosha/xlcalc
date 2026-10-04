/**
 * Разбор формул: токенизатор, AST, проверка с кодами, канонический текст. От движка не зависит.
 *
 * Хранимый текст формулы без `=`. Канон: имена функций заглавными, лист всегда в кавычках
 * (`'Данные'!A1`), строки с Excel-удвоением `""`, без пробелов.
 */
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH, type CheckCode, type ErrorCode, type NameCheck, type ValidateContext, type ValidateResult } from './contract'
import { FUNCTIONS } from './functions'
import { colToIndex, MAX_COLS, MAX_ROWS } from './model'

export interface RangeRef {
  /** Строки и колонки с нуля; у целой колонки r1 = 0, r2 = MAX_ROWS - 1. */
  r1: number
  c1: number
  r2: number
  c2: number
  kind: 'cell' | 'range' | 'cols' | 'rows'
  /** Исходный текст ссылки без листа (`$D$2:$D$3401`), для канона. */
  text: string
}

export type Ast =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; code: ErrorCode }
  | { k: 'ref'; sheet: string | null; ref: RangeRef }
  | { k: 'name'; name: string }
  | { k: 'fn'; name: string; args: Ast[] }
  | { k: 'un'; op: '-' | '+'; a: Ast }
  | { k: 'pct'; a: Ast }
  | { k: 'bin'; op: BinOp; a: Ast; b: Ast }
  | { k: 'paren'; a: Ast }

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>='

export class FormulaError extends Error {
  constructor(
    readonly code: CheckCode,
    message: string,
    readonly at: number,
  ) {
    super(message)
  }
}

const ERROR_LITERALS: Record<string, ErrorCode> = {
  '#DIV/0!': 'DIV0',
  '#N/A': 'NA',
  '#NAME?': 'NAME',
  '#NUM!': 'NUM',
  '#REF!': 'REF',
  '#VALUE!': 'VALUE',
  '#NULL!': 'ERROR',
  '#SPILL!': 'SPILL',
}

export const ERROR_TEXT: Record<ErrorCode, string> = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  CYCLE: '#CYCLE!',
  SPILL: '#SPILL!',
  ERROR: '#ERROR!',
}

// ---------- токенизатор ----------

type Tok =
  | { t: 'num'; v: number; at: number }
  | { t: 'str'; v: string; at: number }
  | { t: 'err'; code: ErrorCode; at: number }
  | { t: 'ref'; sheet: string | null; ref: RangeRef; at: number }
  | { t: 'ident'; v: string; at: number }
  | { t: 'fn'; v: string; at: number }
  | { t: 'op'; v: string; at: number }
  | { t: '('; at: number }
  | { t: ')'; at: number }
  | { t: ','; at: number }
  | { t: 'eof'; at: number }

const IDENT = /^[\p{L}_\\][\p{L}\p{N}_.]*/u
const QUOTED_SHEET = /^'((?:[^']|'')+)'!/
const UNQUOTED_SHEET = /^([\p{L}_][\p{L}\p{N}_.]*)!/u
const CELL_RANGE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]+))?/
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/
const ROW_RANGE = /^(\$?)([0-9]+):(\$?)([0-9]+)/
const NUMBER = /^(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/
const IDENT_CHAR = /[\p{L}\p{N}_.(]/u

function bad(at: number, message: string, code: CheckCode = 'syntax'): never {
  throw new FormulaError(code, message, at)
}

function rowNo(s: string, at: number): number {
  const n = Number(s)
  if (n < 1 || n > MAX_ROWS) bad(at, `row ${s} out of range`)
  return n - 1
}

function colNo(s: string, at: number): number {
  const n = colToIndex(s)
  if (n >= MAX_COLS) bad(at, `column ${s} out of range`)
  return n
}

/** Ссылка без листа с позиции `at` в `rest`, или null. */
function readRef(rest: string, at: number): { ref: RangeRef; len: number } | null {
  const next = (len: number) => rest[len] ?? ''
  let m = CELL_RANGE.exec(rest)
  if (m && !IDENT_CHAR.test(next(m[0].length))) {
    const r1 = rowNo(m[4], at)
    const c1 = colNo(m[2], at)
    const two = m[6] !== undefined
    const r2 = two ? rowNo(m[8], at) : r1
    const c2 = two ? colNo(m[6], at) : c1
    return {
      ref: { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2), kind: two ? 'range' : 'cell', text: m[0].toUpperCase() },
      len: m[0].length,
    }
  }
  m = COL_RANGE.exec(rest)
  if (m && !IDENT_CHAR.test(next(m[0].length))) {
    const c1 = colNo(m[2], at)
    const c2 = colNo(m[4], at)
    return { ref: { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(c1, c2), c2: Math.max(c1, c2), kind: 'cols', text: m[0].toUpperCase() }, len: m[0].length }
  }
  m = ROW_RANGE.exec(rest)
  if (m && !IDENT_CHAR.test(next(m[0].length))) {
    const r1 = rowNo(m[2], at)
    const r2 = rowNo(m[4], at)
    return { ref: { c1: 0, c2: MAX_COLS - 1, r1: Math.min(r1, r2), r2: Math.max(r1, r2), kind: 'rows', text: m[0] }, len: m[0].length }
  }
  return null
}

const OPS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%']

function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++
      continue
    }
    const rest = src.slice(i)
    if (ch === '"') {
      let j = i + 1
      let v = ''
      for (;;) {
        if (j >= src.length) bad(i, 'unterminated string')
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            v += '"'
            j += 2
            continue
          }
          break
        }
        v += src[j++]
      }
      out.push({ t: 'str', v, at: i })
      i = j + 1
      continue
    }
    if (ch === '#') {
      const lit = Object.keys(ERROR_LITERALS).find((e) => rest.toUpperCase().startsWith(e))
      if (!lit) bad(i, 'unknown error literal')
      out.push({ t: 'err', code: ERROR_LITERALS[lit], at: i })
      i += lit.length
      continue
    }
    if (ch === '[') bad(i, 'external workbook reference', 'external_ref')
    if (ch === '{') bad(i, 'array constants are not supported')
    if (ch === '(' || ch === ')' || ch === ',') {
      out.push({ t: ch, at: i } as Tok)
      i++
      continue
    }
    // лист с кавычками или без, потом обязательно ссылка
    let sheet: string | null = null
    let sm = QUOTED_SHEET.exec(rest)
    if (sm) sheet = sm[1].replaceAll("''", "'")
    else if ((sm = UNQUOTED_SHEET.exec(rest))) sheet = sm[1]
    if (sm) {
      const r = readRef(src.slice(i + sm[0].length), i)
      if (!r) bad(i + sm[0].length, 'reference expected after sheet name')
      out.push({ t: 'ref', sheet, ref: r.ref, at: i })
      i += sm[0].length + r.len
      continue
    }
    if (ch === "'") bad(i, 'quoted sheet name without reference')
    if (/[0-9$A-Za-z]/.test(ch)) {
      const r = readRef(rest, i)
      if (r) {
        out.push({ t: 'ref', sheet: null, ref: r.ref, at: i })
        i += r.len
        continue
      }
    }
    const nm = NUMBER.exec(rest)
    if (nm) {
      out.push({ t: 'num', v: Number(nm[0]), at: i })
      i += nm[0].length
      continue
    }
    const im = IDENT.exec(rest)
    if (im) {
      let j = i + im[0].length
      while (src[j] === ' ') j++
      if (src[j] === '(') out.push({ t: 'fn', v: im[0].toUpperCase(), at: i })
      else out.push({ t: 'ident', v: im[0], at: i })
      i += im[0].length
      continue
    }
    const op = OPS.find((o) => rest.startsWith(o))
    if (op) {
      out.push({ t: 'op', v: op, at: i })
      i += op.length
      continue
    }
    bad(i, `unexpected character ${JSON.stringify(ch)}`)
  }
  out.push({ t: 'eof', at: src.length })
  return out
}

// ---------- разбор (рекурсивный спуск) ----------

const LEVELS: string[][] = [['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/'], ['^']]

export function parse(src: string): Ast {
  if (src.trim() === '') bad(0, 'empty formula', 'empty')
  if (src.length > MAX_FORMULA_LENGTH) bad(MAX_FORMULA_LENGTH, `formula longer than ${MAX_FORMULA_LENGTH}`, 'too_long')
  const toks = tokenize(src)
  let p = 0
  let depth = 0
  const peek = () => toks[p]
  const enter = (at: number) => {
    if (++depth > MAX_FORMULA_DEPTH) bad(at, `nesting deeper than ${MAX_FORMULA_DEPTH}`, 'too_deep')
  }

  const binary = (level: number): Ast => {
    if (level === LEVELS.length) return unary()
    let a = binary(level + 1)
    for (;;) {
      const t = peek()
      if (t.t !== 'op' || !LEVELS[level].includes(t.v)) return a
      p++
      const b = binary(level + 1)
      a = { k: 'bin', op: t.v as BinOp, a, b }
    }
  }
  // в Excel унарный минус сильнее степени: -2^2 = 4
  const unary = (): Ast => {
    const t = peek()
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) {
      p++
      enter(t.at)
      const a = unary()
      depth--
      return { k: 'un', op: t.v, a }
    }
    return postfix()
  }
  const postfix = (): Ast => {
    let a = primary()
    while (peek().t === 'op' && (peek() as { v: string }).v === '%') {
      p++
      a = { k: 'pct', a }
    }
    return a
  }
  const primary = (): Ast => {
    const t = toks[p++]
    switch (t.t) {
      case 'num':
        return { k: 'num', v: t.v }
      case 'str':
        return { k: 'str', v: t.v }
      case 'err':
        return { k: 'err', code: t.code }
      case 'ref':
        return { k: 'ref', sheet: t.sheet, ref: t.ref }
      case 'ident': {
        const up = t.v.toUpperCase()
        if (up === 'TRUE' || up === 'FALSE') return { k: 'bool', v: up === 'TRUE' }
        return { k: 'name', name: t.v }
      }
      case '(': {
        enter(t.at)
        const a = binary(0)
        depth--
        expect(')')
        return { k: 'paren', a }
      }
      case 'fn': {
        expect('(')
        enter(t.at)
        const args: Ast[] = []
        if (peek().t !== ')') {
          for (;;) {
            args.push(binary(0))
            if (peek().t === ',') {
              p++
              continue
            }
            break
          }
        }
        depth--
        expect(')')
        return { k: 'fn', name: t.v, args }
      }
      default:
        return bad(t.at, t.t === 'eof' ? 'unexpected end of formula' : `unexpected ${describe(t)}`)
    }
  }
  const expect = (k: '(' | ')') => {
    const t = toks[p]
    if (t.t !== k) bad(t.at, `expected ${k}, found ${describe(t)}`)
    p++
  }

  const ast = binary(0)
  const t = peek()
  if (t.t !== 'eof') bad(t.at, t.t === ',' ? 'comma outside of a function call' : `unexpected ${describe(t)}`)
  return ast
}

function describe(t: Tok): string {
  switch (t.t) {
    case 'op':
    case 'ident':
    case 'fn':
      return JSON.stringify(t.v)
    case 'eof':
      return 'end of formula'
    default:
      return t.t
  }
}

// ---------- обход и печать ----------

export function walk(ast: Ast, visit: (n: Ast) => void): void {
  visit(ast)
  switch (ast.k) {
    case 'fn':
      for (const a of ast.args) walk(a, visit)
      break
    case 'un':
    case 'pct':
    case 'paren':
      walk(ast.a, visit)
      break
    case 'bin':
      walk(ast.a, visit)
      walk(ast.b, visit)
      break
  }
}

/** Листы сравниваются без учёта регистра, как в Excel. */
export const sheetKey = (s: string) => s.toLocaleUpperCase('ru-RU')

export const quoteSheet = (s: string) => `'${s.replaceAll("'", "''")}'`

export interface PrintOptions {
  str?: (s: string) => string
  bool?: (b: boolean) => string
  /** Имя листа в написании книги. */
  sheetOf?: (s: string) => string
}

/** Печать AST; по умолчанию канон. Другие движки (замер в `tools/probes/wb1`) подставляют свои литералы. */
export function printFormula(ast: Ast, opts: PrintOptions = {}): string {
  const str = opts.str ?? ((s: string) => `"${s.replaceAll('"', '""')}"`)
  const bool = opts.bool ?? ((b: boolean) => (b ? 'TRUE' : 'FALSE'))
  const sheetOf = opts.sheetOf ?? ((s: string) => s)
  const go = (n: Ast): string => {
    switch (n.k) {
      case 'num':
        return String(n.v)
      case 'str':
        return str(n.v)
      case 'bool':
        return bool(n.v)
      case 'err':
        return ERROR_TEXT[n.code]
      case 'ref':
        return n.sheet === null ? n.ref.text : `${quoteSheet(sheetOf(n.sheet))}!${n.ref.text}`
      case 'name':
        return n.name
      case 'fn':
        return `${n.name}(${n.args.map(go).join(',')})`
      case 'un':
        return n.op + go(n.a)
      case 'pct':
        return go(n.a) + '%'
      case 'bin':
        return go(n.a) + n.op + go(n.b)
      case 'paren':
        return `(${go(n.a)})`
    }
  }
  return go(ast)
}

/** Канонический хранимый текст (без `=`). */
export const canonical = (ast: Ast, sheetOf?: (s: string) => string) => printFormula(ast, { sheetOf })

// ---------- проверка ----------

/** Узел-диапазон больше одной ячейки (имя считается по `isRangeName`). */
function isArea(n: Ast, isRangeName: (name: string) => boolean | undefined): boolean {
  if (n.k === 'paren') return isArea(n.a, isRangeName)
  if (n.k === 'ref') return n.ref.kind !== 'cell' && !(n.ref.r1 === n.ref.r2 && n.ref.c1 === n.ref.c2)
  if (n.k === 'name') return isRangeName(n.name) === true
  return false
}

/**
 * Форма аргументов: число аргументов функции и диапазоны только там, где функция их ждёт.
 * Диапазон в скалярном месте значит вычисление над массивом: в этой версии `unsupported`.
 */
export function checkShape(ast: Ast, src: string, isRangeName: (name: string) => boolean | undefined, external: (fn: string) => boolean): FormulaError | null {
  const at = (needle: string) => Math.max(0, src.toUpperCase().indexOf(needle.toUpperCase()))
  const visit = (n: Ast, rangeOk: boolean): FormulaError | null => {
    switch (n.k) {
      case 'ref':
      case 'name':
        return !rangeOk && isArea(n, isRangeName) ? new FormulaError('unsupported', 'a range is used where a single value is expected (array arithmetic is not supported)', at(n.k === 'name' ? n.name : n.ref.text)) : null
      case 'paren':
        return visit(n.a, rangeOk)
      case 'un':
      case 'pct':
        return visit(n.a, false)
      case 'bin':
        return visit(n.a, false) ?? visit(n.b, false)
      case 'fn': {
        if (external(n.name)) {
          for (const a of n.args) {
            const lit = a.k === 'num' || a.k === 'str' || a.k === 'bool' || (a.k === 'un' && a.a.k === 'num')
            if (!lit) return new FormulaError('unsupported', `arguments of ${n.name} must be literals`, at(n.name))
          }
          return null
        }
        const spec = FUNCTIONS[n.name]
        if (!spec) return null
        const k = n.args.length
        if (k < spec.min || k > spec.max || (spec.parity === 'odd' && k % 2 === 0) || (spec.parity === 'even' && k % 2 === 1)) {
          const count = spec.min === spec.max ? `${spec.min}` : `${spec.min} to ${spec.max}${spec.parity ? ` (${spec.parity})` : ''}`
          return new FormulaError('syntax', `${n.name} takes ${count} arguments, got ${k}`, at(n.name))
        }
        for (let i = 0; i < k; i++) {
          const e = visit(n.args[i], spec.range(i))
          if (e) return e
        }
        return null
      }
      default:
        return null
    }
  }
  return visit(ast, false)
}

/** Проверка формулы: разбор, функции из белого списка или внешние, листы, имена, форма аргументов. */
export function validate(text: string, ctx: ValidateContext = {}, isRangeName: (name: string) => boolean | undefined = () => undefined): ({ ok: true; ast: Ast } & Extract<ValidateResult, { ok: true }>) | Extract<ValidateResult, { ok: false }> {
  let ast: Ast
  try {
    ast = parse(text)
  } catch (e) {
    if (e instanceof FormulaError) return { ok: false, code: e.code, message: e.message, at: e.at }
    throw e
  }
  const externalSet = new Set((ctx.externalFunctions ?? []).map((f) => f.toUpperCase()))
  const functions = new Set<string>()
  const sheets = new Set<string>()
  const names = new Set<string>()
  const externals: string[] = []
  walk(ast, (n) => {
    if (n.k === 'fn') {
      functions.add(n.name)
      if (externalSet.has(n.name) && !FUNCTIONS[n.name]) externals.push(canonical(n))
    } else if (n.k === 'ref' && n.sheet !== null) sheets.add(n.sheet)
    else if (n.k === 'name') names.add(n.name)
  })
  const fail = (code: CheckCode, message: string, needle: string) => ({ ok: false as const, code, message, at: Math.max(0, text.toUpperCase().indexOf(needle.toUpperCase())) })
  for (const f of functions) if (!FUNCTIONS[f] && !externalSet.has(f)) return fail('unknown_function', `unknown function ${f}`, f)
  if (ctx.sheets) {
    const known = new Set(ctx.sheets.map(sheetKey))
    for (const s of sheets) if (!known.has(sheetKey(s))) return fail('unknown_sheet', `unknown sheet ${s}`, s)
  }
  if (ctx.names) {
    const known = new Set(ctx.names.map((n) => n.toUpperCase()))
    for (const n of names) if (!known.has(n.toUpperCase())) return fail('unknown_name', `unknown name ${n}`, n)
  }
  const shape = checkShape(ast, text, isRangeName, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (shape) return { ok: false, code: shape.code, message: shape.message, at: shape.at }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)] }
}

const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]{0,254}$/
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}[0-9]+$|^[Rr][0-9]*[Cc][0-9]*$/

/** Допустимое имя диапазона: латиница, не адрес ячейки, не R1C1, не TRUE/FALSE. */
export function validateName(name: string): NameCheck {
  if (!NAME_RE.test(name)) return { ok: false, message: 'a name must start with a latin letter or _ and contain only latin letters, digits, _ and .' }
  if (LOOKS_LIKE_CELL.test(name)) return { ok: false, message: 'a name must not look like a cell address' }
  if (/^(TRUE|FALSE)$/i.test(name)) return { ok: false, message: 'TRUE and FALSE are reserved' }
  return { ok: true }
}
