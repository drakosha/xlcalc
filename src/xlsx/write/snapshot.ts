/**
 * What `writeXlsx` is given: a snapshot `{ book, result? }`, the options, and which values of formulas the file stores
 * (`values`). The result is `compute`'s, a session's
 * (`getSnapshot`) or the values a file stored (`readXlsx`'s `cached`); with `values: 'require'` every formula and every
 * cell an array fills has its value in it, as a result of the book has.
 */
import type { BookInput, CellResult } from '../../contract.js'
import { checkBook } from '../../input.js'
import { InputError, what } from '../../input-error.js'
import { object, onlyKeys } from '../../input/basic.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { sheetKey } from '../../syntax/print.js'
import { optionalLocale } from '../../locale/load.js'
import type { Locale } from '../../locale/types.js'
import type { WriteXlsxOptions, XlsxSnapshot } from '../types.js'

export type Values = NonNullable<WriteXlsxOptions['values']>

/** What the writer reads of a result's cell. */
export type ResultCell = Pick<CellResult, 'value' | 'spill' | 'spillOf' | 'spillFill'>

export interface WriteArgs {
  book: BookInput
  /** The result's cells by sheet of the book, but with `omit`, where the file stores no values of formulas. */
  cells?: Record<string, ResultCell>[]
  /** The addresses of each sheet's cells (`Object.keys`), listed once for the checks and the writer. */
  keys: string[][]
  /** The addresses of each sheet's cells of the result, with `cells`. */
  resultKeys?: string[][]
  mode: 'formulas' | 'values'
  values: Values
  /** `values` was not given and there is no result: the file stores no values of formulas, which is reported. */
  unasked: boolean
  /** The language of the issues (`opts.locale`), en-US without one. */
  locale?: Locale
}

const RESULT = "a result: compute's, a session's getSnapshot().result or readXlsx's cached, { sheets: [{ name, cells }] }"
const SESSION = 'a session is not a snapshot: pass session.getSnapshot() (awaited for an asynchronous session)'
const RESULT_ALONE = 'a result is not a snapshot: pass { book, result } with the workbook it was calculated from'
const BOOK_ALONE = ' (a workbook is not a snapshot: pass { book, result })'
const PRESERVE = '(values: "preserve" writes the values it has)'

/** A session or its proxy, which a snapshot is taken of. */
const isSession = (v: unknown) => typeof (v as { getSnapshot?: unknown } | null)?.getSnapshot === 'function'

/** A result (`compute`'s, a session's `getResult()`), told by its hashes: it has sheets, and a book does too. */
const isResult = (v: unknown) => typeof (v as { hash?: unknown } | null)?.hash === 'string' && typeof (v as { inputHash?: unknown }).inputHash === 'string'

/** The result's cells by sheet with their addresses; the result must be of a result's shape and the book's (the same sheets in order). */
function resultCells(book: BookInput, result: unknown, path: string): { cells: Record<string, ResultCell>; keys: string[] }[] {
  if (isSession(result)) throw new InputError(path, SESSION)
  const r = object(result, path, RESULT)
  if (!Array.isArray(r.sheets)) throw new InputError(path, `expected ${RESULT}, got an object without sheets`)
  const sheets = r.sheets as { name?: unknown; cells?: unknown }[]
  if (sheets.length !== book.sheets.length || sheets.some((s, i) => typeof s?.name !== 'string' || sheetKey(s.name) !== sheetKey(book.sheets[i].name)))
    throw new InputError(path, "expected the result of this workbook: its sheets differ from the workbook's")
  return sheets.map((s, i) => {
    const p = `${path}.sheets[${i}].cells`
    const cells = object(s.cells, p, "the result's cells by address, such as { B1: { value: 200 } }")
    const keys = Object.keys(cells)
    for (const a of keys) {
      const c = cells[a]
      if (typeof c !== 'object' || c === null || !('value' in c)) throw new InputError(`${p}[${JSON.stringify(a)}]`, `expected a cell of a result, such as { value: 200 }, got ${what(c)}`)
    }
    return { cells: cells as Record<string, ResultCell>, keys }
  })
}

/** The arguments of `writeXlsx`, checked: a snapshot and options. */
export function writeArgs(snapshot: XlsxSnapshot, opts: unknown): WriteArgs {
  if (isSession(snapshot)) throw new InputError('snapshot', SESSION)
  if (isResult(snapshot)) throw new InputError('snapshot', RESULT_ALONE)
  const given = object(snapshot, 'snapshot', 'a snapshot { book, result }')
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { values: 'preserve' }")
  onlyKeys(given, 'snapshot', ['book', 'result'], 'cached' in given ? ' (a file read: pass { book: read.book, result: read.cached })' : 'sheets' in given ? BOOK_ALONE : '')
  onlyKeys(o, 'opts', ['mode', 'values', 'locale'], 'result' in o ? ' (the result goes in the snapshot: writeXlsx({ book, result }))' : '')
  // a sheet without looks is written as given, its entries in their order; one with looks as the engine takes it
  const listed = new Map<object, string[]>()
  const book = checkBook(given.book, 'snapshot.book', true, listed)
  // a sheet checkBook gave anew (formulas without their `=`, looks alone taken out) is listed here
  const keys = book.sheets.map((s) => listed.get(s.cells) ?? Object.keys(s.cells))
  const { result } = given
  const path = 'snapshot.result'
  const mode = o.mode ?? 'formulas'
  if (mode !== 'formulas' && mode !== 'values') throw new InputError('opts.mode', `expected "formulas" or "values", got ${JSON.stringify(mode)}`)
  const has = result !== undefined
  const values = o.values ?? (has ? 'require' : 'omit')
  if (values !== 'require' && values !== 'omit' && values !== 'preserve') throw new InputError('opts.values', `expected "require", "omit" or "preserve", got ${JSON.stringify(values)}`)
  if (mode === 'values' && !has) throw new InputError(path, `the values mode takes the values from the result: pass { book, result }`)
  if (mode === 'values' && values === 'omit') throw new InputError('opts.values', `the values mode writes values only: "omit" would leave formula cells empty`)
  if (values !== 'omit' && !has) throw new InputError(path, `values: "${values}" takes the values of formulas from the result: pass { book, result }`)
  const res = has ? resultCells(book, result, path) : undefined
  const cells = res?.map((r) => r.cells)
  const resultKeys = res?.map((r) => r.keys)
  if (values === 'require') requireValues(book, keys, cells!, resultKeys!, path)
  const locale = optionalLocale(o.locale, 'opts.locale')
  return { book, keys, mode, values, unasked: o.values === undefined && !has, ...(values !== 'omit' ? { cells, resultKeys } : {}), ...(locale ? { locale } : {}) }
}

/** Every formula of the book and every cell of an array's range (but past a large array's `spillFill`) has its value. */
function requireValues(book: BookInput, keys: string[][], cells: Record<string, ResultCell>[], resultKeys: string[][], path: string): void {
  const missing = (where: string) => new InputError(path, `expected the result of this workbook: ${where} has no value in the result ${PRESERVE}`)
  book.sheets.forEach((s, i) => {
    const r = cells[i]
    for (const a of keys[i]) if (s.cells[a].formula !== undefined && r[a] === undefined) throw missing(`the formula of '${s.name}'!${a}`)
    for (const a of resultKeys[i]) {
      const x = r[a].spillFill === undefined && r[a].spill !== undefined ? parseRange(r[a].spill!) : null
      if (!x) continue
      for (let row = x.r1; row <= x.r2; row++)
        for (let c = x.c1; c <= x.c2; c++) if (r[formatAddress(row, c)] === undefined) throw missing(`'${s.name}'!${formatAddress(row, c)} of the array of ${a}`)
    }
  })
}
