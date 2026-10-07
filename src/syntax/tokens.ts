/** The tokenizer: references (sheets, ranges across sheets, spill references, lost ones), numbers, strings, error literals, identifiers (names with a sheet too), operators. */
import type { ErrorCode } from '../contract.js'
import { colToIndex, MAX_COLS, MAX_ROWS, readNumber } from '../model.js'
import { bad, type RangeRef } from './ast.js'

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

export type Tok =
  | { t: 'num'; v: number; at: number }
  | { t: 'str'; v: string; at: number }
  | { t: 'err'; code: ErrorCode; at: number; sheet?: string; ref?: string }
  /** `last`: the last sheet of a reference across sheets (`Sheet1:Sheet3!A1`). */
  | { t: 'ref'; sheet: string | null; last?: string; ref: RangeRef; at: number }
  /** `sheet`: a name seen from that sheet (`Sheet2!rate`). */
  | { t: 'ident'; v: string; at: number; sheet?: string }
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
/** Two sheets, the last quoted or not: Excel saves `Jan:'Mar'!A1` too. */
const UNQUOTED_SHEETS = /^([\p{L}_][\p{L}\p{N}_.]*):(?:([\p{L}_][\p{L}\p{N}_.]*)|'((?:[^']|'')+)')!/u
const CELL_RANGE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]+))?/
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/
const ROW_RANGE = /^(\$?)([0-9]+):(\$?)([0-9]+)/
const NUMBER = /^(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/
const IDENT_CHAR = /[\p{L}\p{N}_.(]/u

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

/** The length of a reference read at `at`, marking it a spill reference when `#` follows a cell (`E2#`). */
function spilled(r: { ref: RangeRef; len: number }, next: string | undefined, at: number): number {
  if (next !== '#') return r.len
  if (r.ref.kind !== 'cell') bad(at, 'a spill reference (#) follows a single cell')
  r.ref.spill = true
  return r.len + 1
}

const OPS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '@']

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
      // a reference whose sheet was deleted: `#REF!A1`, `#REF!#REF!`; a lost spill reference ends in `#`
      let tail = lit === '#REF!' ? (readRef(src.slice(i + 5), i)?.ref.text ?? (rest.slice(5).toUpperCase().startsWith('#REF!') ? '#REF!' : '')) : ''
      if (lit === '#REF!' && src[i + 5 + tail.length] === '#') tail += '#'
      out.push({ t: 'err', code: ERROR_LITERALS[lit], at: i, ...(tail ? { ref: tail } : {}) })
      i += lit.length + tail.length
      continue
    }
    if (ch === '[') bad(i, 'external workbook reference', 'external_ref')
    if (ch === '(' || ch === ')' || ch === ',' || ch === '{' || ch === '}' || ch === ';') {
      out.push({ t: ch, at: i } as Tok)
      i++
      continue
    }
    // a sheet, quoted or not, or two sheets (no sheet name has `:`), must be followed by a reference or a name
    let sheet: string | null = null
    let last: string | undefined
    let sm = QUOTED_SHEET.exec(rest)
    if (sm) {
      sheet = sm[1].replaceAll("''", "'")
      const k = sheet.indexOf(':')
      if (k >= 0) [sheet, last] = [sheet.slice(0, k), sheet.slice(k + 1)]
    } else if (rest.includes('!') && (sm = UNQUOTED_SHEETS.exec(rest))) [sheet, last] = [sm[1], sm[2] ?? sm[3].replaceAll("''", "'")]
    else if ((sm = UNQUOTED_SHEET.exec(rest))) sheet = sm[1]
    if (sm && last === undefined && src.slice(i + sm[0].length, i + sm[0].length + 5).toUpperCase() === '#REF!') {
      // a reference Excel lost on a sheet that is still there: `'S'!#REF!`, `'S'!#REF!#`
      const spill = src[i + sm[0].length + 5] === '#'
      out.push({ t: 'err', code: 'REF', sheet: sheet!, at: i, ...(spill ? { ref: '#' } : {}) })
      i += sm[0].length + 5 + (spill ? 1 : 0)
      continue
    }
    if (sm) {
      const after = src.slice(i + sm[0].length)
      const r = readRef(after, i)
      if (r) {
        out.push({ t: 'ref', sheet, ...(last !== undefined ? { last } : {}), ref: r.ref, at: i })
        if (last !== undefined && after[r.len] === '#') bad(i, 'a spill reference (#) follows a single cell of one sheet')
        i += sm[0].length + spilled(r, after[r.len], i)
        continue
      }
      // a name as one sheet sees it, not a function
      const nm = last === undefined ? IDENT.exec(after) : null
      if (!nm || after[nm[0].length] === '(') bad(i + sm[0].length, 'reference expected after sheet name')
      out.push({ t: 'ident', v: nm[0], at: i, sheet: sheet! })
      i += sm[0].length + nm[0].length
      continue
    }
    if (ch === "'") bad(i, 'quoted sheet name without reference')
    if (/[0-9$A-Za-z]/.test(ch)) {
      const r = readRef(rest, i)
      if (r) {
        out.push({ t: 'ref', sheet: null, ref: r.ref, at: i })
        i += spilled(r, src[i + r.len], i)
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
