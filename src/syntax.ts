/**
 * Formula parsing: tokenizer, AST, checks with codes, canonical text. Independent of the engine.
 *
 * The stored formula text has no `=`. Canonical form: function names upper-case, the sheet always quoted
 * (`'Data'!A1`), strings with Excel `""` doubling, no whitespace.
 */
import { shapes } from './arrays/shape.js'
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH, type CheckCode, type ErrorCode, type NameCheck, type ValidateContext, type ValidateResult } from './contract.js'
import { FUNCTIONS } from './functions.js'
import { colToIndex, MAX_COLS, MAX_ROWS, readNumber } from './model.js'

export interface RangeRef {
  /** Zero-based rows and columns; a whole column has r1 = 0, r2 = MAX_ROWS - 1. */
  r1: number
  c1: number
  r2: number
  c2: number
  kind: 'cell' | 'range' | 'cols' | 'rows'
  /** Source text of the reference without the sheet (`$D$2:$D$3401`), for the canonical form. */
  text: string
}

export type Ast =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  /** `sheet` and `ref` keep what Excel writes around a reference it lost: `'S'!#REF!`, `#REF!A1`, `#REF!#REF!`. */
  | { k: 'err'; code: ErrorCode; sheet?: string; ref?: string }
  | { k: 'ref'; sheet: string | null; ref: RangeRef }
  | { k: 'name'; name: string }
  | { k: 'fn'; name: string; args: Ast[] }
  | { k: 'un'; op: '-' | '+'; a: Ast }
  | { k: 'pct'; a: Ast }
  | { k: 'bin'; op: BinOp; a: Ast; b: Ast }
  | { k: 'paren'; a: Ast }
  /** An array constant `{1,2;3,4}`: rows of numbers, text, booleans and errors. */
  | { k: 'arr'; rows: Lit[][] }

/** A value of an array constant. */
export type Lit = Extract<Ast, { k: 'num' | 'str' | 'bool' | 'err' }>

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
  CALC: '#CALC!',
  ERROR: '#ERROR!',
}

// ---------- tokenizer ----------

export type Tok =
  | { t: 'num'; v: number; at: number }
  | { t: 'str'; v: string; at: number }
  | { t: 'err'; code: ErrorCode; at: number; sheet?: string; ref?: string }
  | { t: 'ref'; sheet: string | null; ref: RangeRef; at: number }
  | { t: 'ident'; v: string; at: number }
  | { t: 'fn'; v: string; at: number }
  | { t: 'op'; v: string; at: number }
  | { t: '('; at: number }
  | { t: ')'; at: number }
  | { t: ','; at: number }
  | { t: '{' | '}' | ';'; at: number }
  | { t: 'eof'; at: number }

/** An identifier (function or name): a letter, `_` or `\`, then letters, digits, `_` and `.`; the locale lexer reads the same. */
export const IDENT_RULE = '[\\p{L}_\\\\][\\p{L}\\p{N}_.]*'
const IDENT = new RegExp(`^${IDENT_RULE}`, 'u')
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

/** A reference without a sheet at position `at` in `rest`, or null. */
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

export function tokenize(src: string): Tok[] {
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
      // a reference whose sheet was deleted: `#REF!A1`, `#REF!#REF!`
      const tail = lit === '#REF!' ? (readRef(src.slice(i + 5), i)?.ref.text ?? (rest.slice(5).toUpperCase().startsWith('#REF!') ? '#REF!' : undefined)) : undefined
      out.push({ t: 'err', code: ERROR_LITERALS[lit], at: i, ...(tail ? { ref: tail } : {}) })
      i += lit.length + (tail?.length ?? 0)
      continue
    }
    if (ch === '[') bad(i, 'external workbook reference', 'external_ref')
    if (ch === '(' || ch === ')' || ch === ',' || ch === '{' || ch === '}' || ch === ';') {
      out.push({ t: ch, at: i } as Tok)
      i++
      continue
    }
    // a sheet, quoted or not, must be followed by a reference
    let sheet: string | null = null
    let sm = QUOTED_SHEET.exec(rest)
    if (sm) sheet = sm[1].replaceAll("''", "'")
    else if ((sm = UNQUOTED_SHEET.exec(rest))) sheet = sm[1]
    if (sm && src.slice(i + sm[0].length, i + sm[0].length + 5).toUpperCase() === '#REF!') {
      // a reference Excel lost on a sheet that is still there: `'S'!#REF!`
      out.push({ t: 'err', code: 'REF', sheet: sheet!, at: i })
      i += sm[0].length + 5
      continue
    }
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
      out.push({ t: 'num', v: readNumber(nm[0]), at: i })
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

// ---------- parsing (recursive descent) ----------

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
  // in Excel unary minus binds tighter than ^: -2^2 = 4
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
        return { k: 'err', code: t.code, ...(t.sheet !== undefined ? { sheet: t.sheet } : {}), ...(t.ref !== undefined ? { ref: t.ref } : {}) }
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
      case '{':
        return constant(t.at)
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
  // rows separated by `;`, values by `,`, all rows of one length, as Excel requires
  const constant = (at: number): Ast => {
    const rows: Lit[][] = [[]]
    for (;;) {
      rows.at(-1)!.push(literal())
      const t = toks[p++]
      if (t.t === ',') continue
      if (t.t === ';') {
        rows.push([])
        continue
      }
      if (t.t !== '}') bad(t.at, `expected , ; or } in an array constant, found ${describe(t)}`)
      break
    }
    if (rows.some((r) => r.length !== rows[0].length)) bad(at, 'the rows of an array constant must have the same number of values')
    return { k: 'arr', rows }
  }
  const literal = (): Lit => {
    const t = toks[p++]
    if (t.t === 'op' && (t.v === '-' || t.v === '+') && toks[p].t === 'num') {
      const n = toks[p++] as { v: number }
      return { k: 'num', v: t.v === '-' ? -n.v : n.v }
    }
    if (t.t === 'num') return { k: 'num', v: t.v }
    if (t.t === 'str') return { k: 'str', v: t.v }
    if (t.t === 'err' && t.sheet === undefined && t.ref === undefined) return { k: 'err', code: t.code }
    if (t.t === 'ident' && /^(TRUE|FALSE)$/i.test(t.v)) return { k: 'bool', v: t.v.toUpperCase() === 'TRUE' }
    return bad(t.at, `an array constant holds numbers, text, booleans and errors only, found ${describe(t)}`)
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

// ---------- traversal and printing ----------

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

/** Sheet names compare case-insensitively, as in Excel. */
export const sheetKey = (s: string) => s.toLocaleUpperCase('ru-RU')

export const quoteSheet = (s: string) => `'${s.replaceAll("'", "''")}'`

export interface PrintOptions {
  str?: (s: string) => string
  bool?: (b: boolean) => string
  /** The sheet name as spelled in the book. */
  sheetOf?: (s: string) => string
}

/** Prints an AST; canonical by default. Adapters for other engines substitute their own literals. */
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
        return (n.sheet === undefined ? '' : `${quoteSheet(sheetOf(n.sheet))}!`) + ERROR_TEXT[n.code] + (n.ref ?? '')
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
      case 'arr':
        return `{${n.rows.map((r) => r.map(go).join(',')).join(';')}}`
    }
  }
  return go(ast)
}

/** Canonical stored text (without `=`). */
export const canonical = (ast: Ast, sheetOf?: (s: string) => string) => printFormula(ast, { sheetOf })

// ---------- checks ----------

/** Why a formula whose result is an array of several values gets `unsupported` (and its cell #SPILL!). */
export const ARRAY_RESULT = 'spilling into neighbouring cells is not supported yet; wrap it in a function that takes an array, such as SUM or INDEX'

/** Argument counts of each function; literal arguments of external functions. */
export function checkCalls(ast: Ast, src: string, external: (fn: string) => boolean): FormulaError | null {
  const at = (needle: string) => Math.max(0, src.toUpperCase().indexOf(needle.toUpperCase()))
  const visit = (n: Ast): FormulaError | null => {
    switch (n.k) {
      case 'paren':
      case 'un':
      case 'pct':
        return visit(n.a)
      case 'bin':
        return visit(n.a) ?? visit(n.b)
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
        for (const a of n.args) {
          const e = visit(a)
          if (e) return e
        }
        return null
      }
      default:
        return null
    }
  }
  return visit(ast)
}

/**
 * Formula check: parsing, whitelisted or external functions, sheets, names, argument counts, and (unless
 * `arrayResult` is false, as in `compute`, which gives such a cell #SPILL! when computed) a result that is
 * an array of several values whatever the data.
 */
export function validate(text: string, ctx: ValidateContext = {}, isRangeName: (name: string) => boolean | undefined = () => undefined, arrayResult = true): ({ ok: true; ast: Ast } & Extract<ValidateResult, { ok: true }>) | Extract<ValidateResult, { ok: false }> {
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
    } else if ((n.k === 'ref' && n.sheet !== null) || (n.k === 'err' && n.sheet !== undefined)) sheets.add(n.sheet!)
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
  const calls = checkCalls(ast, text, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (calls) return { ok: false, code: calls.code, message: calls.message, at: calls.at }
  if (arrayResult && shapes(ast, isRangeName).get(ast) === 2) return { ok: false, code: 'unsupported', message: `the result is an array: ${ARRAY_RESULT}`, at: 0 }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)] }
}

const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]{0,254}$/
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}[0-9]+$|^[Rr][0-9]*[Cc][0-9]*$/

/** A valid range name: Latin letters, not a cell address, not R1C1, not TRUE/FALSE. */
export function validateName(name: string): NameCheck {
  if (!NAME_RE.test(name)) return { ok: false, message: 'a name must start with a latin letter or _ and contain only latin letters, digits, _ and .' }
  if (LOOKS_LIKE_CELL.test(name)) return { ok: false, message: 'a name must not look like a cell address' }
  if (/^(TRUE|FALSE)$/i.test(name)) return { ok: false, message: 'TRUE and FALSE are reserved' }
  return { ok: true }
}
