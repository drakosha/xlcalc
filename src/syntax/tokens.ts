/**
 * The tokenizer: references (sheets, ranges across sheets, spill references, lost ones), structured references to
 * tables, numbers, strings, error literals, identifiers (names with a sheet too), operators. Whitespace between two
 * references (or names, calls, parentheses) is the intersection operator, `' '`; elsewhere it is skipped.
 */
import { PLAIN_SHEET } from '../sheet-pattern.js'
import { wordKey } from '../collate.js'
import type { ErrorCode } from '../contract.js'
import { colToIndex, MAX_COLS, MAX_ROWS, readNumber } from '../model.js'
import { bad, type RangeRef } from './ast.js'
import { ERROR_CODE, FORMULA_ERRORS } from './errors.js'
import { BOOK_SHEET, bracketEnd, callFollows, closingQuote, IDENT_CHAR, IDENT_RULE, isSpace, numberSource } from './scan.js'
import { readSpec, type TableSpec } from './table.js'

export type Tok =
  | { t: 'num'; v: number; at: number }
  | { t: 'str'; v: string; at: number }
  /** `book`: a reference another workbook lost (`'[1]Sheet1'!#REF!`), on its `sheet`. */
  | { t: 'err'; code: ErrorCode; at: number; sheet?: string; ref?: string; book?: number | string }
  /** `last`: the last sheet of a reference across sheets (`Sheet1:Sheet3!A1`). */
  | { t: 'ref'; sheet: string | null; last?: string; ref: RangeRef; at: number }
  /** `sheet`: a name seen from that sheet (`Sheet2!rate`). */
  | { t: 'ident'; v: string; at: number; sheet?: string }
  /**
   * A reference to another workbook's sheet: `book` the number of its link (`[1]Sheet1!A1`, `'[1]My Sheet'!A1`), or the
   * path and file name it is shown with (`[Book2.xlsx]Sheet1!A1`, `'C:\data\[Book2.xlsx]Sheet1'!A1`).
   */
  | { t: 'xref'; book: number | string; sheet: string; ref: RangeRef; at: number }
  /** A defined name of another workbook, of the book (`[1]!Rate`) or of one of its sheets (`[1]Sheet1!Local`). */
  | { t: 'xname'; book: number | string; sheet?: string; v: string; at: number }
  /** A structured reference; `table` null without a table name (`[@Qty]`). */
  | { t: 'table'; table: string | null; spec: TableSpec; at: number }
  | { t: 'fn'; v: string; at: number }
  | { t: 'op'; v: string; at: number }
  | { t: '('; at: number }
  | { t: ')'; at: number }
  | { t: ','; at: number }
  | { t: '{' | '}' | ';'; at: number }
  | { t: 'eof'; at: number }

const IDENT = new RegExp(`^${IDENT_RULE}`, 'u')
const QUOTED_SHEET = /^'((?:[^']|'')+)'!/
const UNQUOTED_SHEET = new RegExp(`^(${PLAIN_SHEET})!`, 'u')
/** Two sheets, the last quoted or not: Excel saves `Jan:'Mar'!A1` too. */
const UNQUOTED_SHEETS = new RegExp(`^(${PLAIN_SHEET}):(?:(${PLAIN_SHEET})|'((?:[^']|'')+)')!`, 'u')
/** A name that reads as a cell (`S1`, `Q4`): Excel does not take it for the first of two sheets, `S1:S3!A1` is `S1:'S3'!A1`. */
const CELL_NAME = /^([A-Za-z]{1,3})([0-9]+)$/
const CELL_RANGE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]+))?/
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/
const ROW_RANGE = /^(\$?)([0-9]+):(\$?)([0-9]+)/
const NUMBER = new RegExp(`^${numberSource('\\.')}`)

function rowNo(s: string, at: number): number {
  const n = Number(s)
  if (n < 1 || n > MAX_ROWS) bad(at, { key: 'rowOutOfRange', row: s })
  return n - 1
}

function colNo(s: string, at: number): number {
  const n = colToIndex(s)
  if (n >= MAX_COLS) bad(at, { key: 'columnOutOfRange', column: s })
  return n
}

/** A reference without a sheet at position `at` in `rest`, or null. */
function readRef(rest: string, at: number): { ref: RangeRef; len: number } | null {
  const next = (len: number) => rest[len] ?? ''
  let m = CELL_RANGE.exec(rest)
  // the second cell of `S1:S3!A1` is a sheet's name, of `A1:E2#` a spill reference: the first is a cell alone, before the range operator
  if (m && m[6] !== undefined && (next(m[0].length) === '!' || next(m[0].length) === '#')) m = CELL_RANGE.exec(rest.slice(0, rest.indexOf(':')))
  // a cell's shape past the sheet, without `$`, is a name, as Excel reads `XFE1` and `A0` (the sets `entry/names-held`,
  // `tables/paste-names-held`)
  if (m && !IDENT_CHAR.test(next(m[0].length)) && !(m[6] === undefined && !m[1] && !m[3] && !inSheet(m[2], m[4]))) {
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
  if (r.ref.kind !== 'cell') bad(at, { key: 'spillAfterRange' })
  r.ref.spill = true
  return r.len + 1
}

const OPS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '@', ':']

/** Whether a column's letters and a row's digits are a cell of the sheet. */
const inSheet = (col: string, row: string) => colToIndex(col) < MAX_COLS && Number(row) >= 1 && Number(row) <= MAX_ROWS

/** Whether a word reads as a cell of the sheet (`S1`, not `XFE1` or `A0`). */
export function readsAsCell(word: string): boolean {
  const m = CELL_NAME.exec(word)
  return !!m && inSheet(m[1], m[2])
}

/** Tokens that end and begin a reference: whitespace between two of them is the intersection operator. */
const ENDS = new Set(['ref', 'xref', 'ident', 'xname', 'table', ')'])
const BEGINS = new Set(['ref', 'xref', 'ident', 'xname', 'table', 'fn', '('])

/** A link's number as a file writes it (`1`), else the path and file name it is shown with. */
const bookOf = (path: string, file: string): number | string => (path === '' && /^[1-9][0-9]*$/.test(file) ? Number(file) : path + file)

/** A quoted sheet with a book in brackets (`'C:\data\[Book2.xlsx]My Sheet'`, `'[1]My Sheet'`): the path, the book, the sheet. */
const BOOK_IN_QUOTES = /^([^[\]]*)\[([^[\]]+)\]([^[\]]*)$/

export function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  // where the last run of whitespace began, if it is just before the token being read
  let space = -1
  const push = (t: Tok) => {
    const last = out.at(-1)
    if (space >= 0 && last && ENDS.has(last.t) && BEGINS.has(t.t)) out.push({ t: 'op', v: ' ', at: space })
    space = -1
    out.push(t)
  }
  /**
   * What follows another workbook's sheet and `!` at `from`: a reference, or a defined name of the book (no sheet) or of
   * the sheet; the end of what was read. A reference across its sheets is not taken (`[1]Jan:Mar!A1`).
   */
  const external = (book: number | string, sheet: string, at: number, from: number): number => {
    if (sheet.includes(':')) bad(at, { key: 'externalAcrossSheets' }, 'external_ref')
    const after = src.slice(from)
    // a reference into it that Excel lost (`[1]Sheet1!#REF!`)
    if (sheet !== '' && after.slice(0, 5).toUpperCase() === '#REF!') {
      push({ t: 'err', code: 'REF', sheet, book, at })
      return from + 5
    }
    const r = sheet === '' ? null : readRef(after, at)
    if (r) {
      if (after[r.len] === '#') bad(at, { key: 'externalSpill' }, 'external_ref')
      push({ t: 'xref', book, sheet, ref: r.ref, at })
      return from + r.len
    }
    const nm = IDENT.exec(after)
    if (!nm || after[nm[0].length] === '(' || after[nm[0].length] === '[') bad(from, { key: 'referenceAfterSheet' })
    push({ t: 'xname', book, ...(sheet === '' ? {} : { sheet }), v: nm[0], at })
    return from + nm[0].length
  }
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (isSpace(ch)) {
      if (space < 0) space = i
      i++
      continue
    }
    const rest = src.slice(i)
    if (ch === '"') {
      const j = closingQuote(src, i)
      if (j < 0) bad(i, { key: 'unterminatedString' })
      push({ t: 'str', v: src.slice(i + 1, j).replaceAll('""', '"'), at: i })
      i = j + 1
      continue
    }
    if (ch === '#') {
      const lit = FORMULA_ERRORS.find((e) => rest.toUpperCase().startsWith(e))
      if (!lit) bad(i, { key: 'unknownErrorLiteral' })
      // a reference whose sheet was deleted: `#REF!A1`, `#REF!#REF!`; a lost spill reference ends in `#`
      let tail = lit === '#REF!' ? (readRef(src.slice(i + 5), i)?.ref.text ?? (rest.slice(5).toUpperCase().startsWith('#REF!') ? '#REF!' : '')) : ''
      if (lit === '#REF!' && src[i + 5 + tail.length] === '#') tail += '#'
      push({ t: 'err', code: ERROR_CODE[lit], at: i, ...(tail ? { ref: tail } : {}) })
      i += lit.length + tail.length
      continue
    }
    if (ch === '[') {
      const close = bracketEnd(src, i)
      const book = BOOK_SHEET.exec(src.slice(close))
      if (book) {
        const sheet = book[0].slice(0, -1)
        i = external(bookOf('', src.slice(i + 1, close - 1)), sheet.startsWith("'") ? sheet.slice(1, -1).replaceAll("''", "'") : sheet, i, close + book[0].length)
        continue
      }
      const { spec, end } = readSpec(src, i, i)
      push({ t: 'table', table: null, spec, at: i })
      i = end
      continue
    }
    if (ch === '(' || ch === ')' || ch === ',' || ch === '{' || ch === '}' || ch === ';') {
      push({ t: ch, at: i } as Tok)
      i++
      continue
    }
    // a sheet, quoted or not, or two sheets (no sheet name has `:`), must be followed by a reference or a name
    let sheet: string | null = null
    let last: string | undefined
    let sm = QUOTED_SHEET.exec(rest)
    const inQuotes = sm && BOOK_IN_QUOTES.exec(sm[1].replaceAll("''", "'"))
    if (sm && inQuotes) {
      i = external(bookOf(inQuotes[1], inQuotes[2]), inQuotes[3], i, i + sm[0].length)
      continue
    }
    if (sm) {
      sheet = sm[1].replaceAll("''", "'")
      // a path, which no sheet's name holds, before another workbook's own name (`'C:\\data\\Book2.xlsx'!Rate`) is no range of sheets
      const k = /[\\/]/.test(sheet) ? -1 : sheet.indexOf(':')
      if (k >= 0) [sheet, last] = [sheet.slice(0, k), sheet.slice(k + 1)]
    } else if (rest.includes('!') && (sm = UNQUOTED_SHEETS.exec(rest)) && !readsAsCell(sm[1])) [sheet, last] = [sm[1], sm[2] ?? sm[3].replaceAll("''", "'")]
    else if ((sm = UNQUOTED_SHEET.exec(rest))) sheet = sm[1]
    if (sm && last === undefined && src.slice(i + sm[0].length, i + sm[0].length + 5).toUpperCase() === '#REF!') {
      // a reference Excel lost on a sheet that is still there: `'S'!#REF!`, `'S'!#REF!#`
      const spill = src[i + sm[0].length + 5] === '#'
      push({ t: 'err', code: 'REF', sheet: sheet!, at: i, ...(spill ? { ref: '#' } : {}) })
      i += sm[0].length + 5 + (spill ? 1 : 0)
      continue
    }
    if (sm) {
      const after = src.slice(i + sm[0].length)
      const r = readRef(after, i)
      if (r) {
        push({ t: 'ref', sheet, ...(last !== undefined ? { last } : {}), ref: r.ref, at: i })
        if (last !== undefined && after[r.len] === '#') bad(i, { key: 'spillAcrossSheets' })
        i += sm[0].length + spilled(r, after[r.len], i)
        continue
      }
      // a name as one sheet sees it, not a function
      const nm = last === undefined ? IDENT.exec(after) : null
      if (!nm || after[nm[0].length] === '(') bad(i + sm[0].length, { key: 'referenceAfterSheet' })
      // a table needs no sheet: Excel drops one written before it
      if (after[nm[0].length] === '[') {
        const start = i + sm[0].length + nm[0].length
        const { spec, end } = readSpec(src, start, i)
        push({ t: 'table', table: nm[0], spec, at: i })
        i = end
        continue
      }
      push({ t: 'ident', v: nm[0], at: i, sheet: sheet! })
      i += sm[0].length + nm[0].length
      continue
    }
    if (ch === "'") bad(i, { key: 'quotedSheetAlone' })
    if (/[0-9$A-Za-z]/.test(ch)) {
      const r = readRef(rest, i)
      if (r) {
        push({ t: 'ref', sheet: null, ref: r.ref, at: i })
        i += spilled(r, src[i + r.len], i)
        continue
      }
    }
    const nm = NUMBER.exec(rest)
    if (nm) {
      push({ t: 'num', v: readNumber(nm[0]), at: i })
      i += nm[0].length
      continue
    }
    const im = IDENT.exec(rest)
    if (im) {
      if (callFollows(src, i + im[0].length)) push({ t: 'fn', v: wordKey(im[0]), at: i })
      else if (src[i + im[0].length] === '[') {
        const { spec, end } = readSpec(src, i + im[0].length, i)
        push({ t: 'table', table: im[0], spec, at: i })
        i = end
        continue
      } else push({ t: 'ident', v: im[0], at: i })
      i += im[0].length
      continue
    }
    const op = OPS.find((o) => rest.startsWith(o))
    if (op) {
      push({ t: 'op', v: op, at: i })
      i += op.length
      continue
    }
    bad(i, { key: 'unexpectedCharacter', char: ch })
  }
  out.push({ t: 'eof', at: src.length })
  return out
}
