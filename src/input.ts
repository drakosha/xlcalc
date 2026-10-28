/**
 * Checks of caller input at the API boundary. Wrong input (a value of the wrong type, a malformed cell
 * address, duplicate sheets, an unknown sheet in an edit) throws `InputError` naming the place and the
 * problem. Formula problems are not input errors: an invalid formula gives an error value in its cell.
 */
import type { BookInput, CachedValues, CalcLimits, CellEdit, CellInput, CellValue, Clock, ComputeOptions, DisplayOptions, EntryOptions, EntryTextOptions, EntryWidthOptions, Execution, ExternalBook, Externals, LayoutEdit, NameInput, NameOptions, NamesContext, SessionOptions, ValidateContext } from './contract.js'
import { overlapping, parseMerge } from './layout/merges.js'
import { checkSpan, SPAN_KEY } from './layout/spans.js'
import { english } from './locale/english.js'
import type { Message } from './locale/messages.js'
import { checkBookLocale, optionalLocale } from './locale/load.js'
import { dateToSerial, MAX_COLS, MAX_ROWS } from './model.js'
import { cornersText, LINES, parseRange, rectText, type Axis } from './structure/area.js'
import { checkArrays } from './array-areas.js'
import { hasLooks, lookAlone, looksTaken } from './looks/sheet.js'
import { ERROR_TEXT, sheetKey } from './syntax.js'
import { checkTables } from './tables/check.js'
import { InputError, what } from './input-error.js'
import { array, checkBoolean, checkNumber, checkOneOf, checkString, isObject, object, onlyKeys, onlyOptions } from './input/basic.js'
import { checkAddress, checkRange } from './input/range.js'
import { checkColor, checkFormat, checkStyleIndex, checkStyles, givenIn } from './input/styles.js'
import { checkSheetView } from './input-view.js'
import { checkSheetProtection } from './input/protection.js'
import { checkConditionalFormats } from './cond/check.js'
import { checkOutline, checkValidations } from './validation/check.js'
import { bookTaken, editsTaken, layoutTaken, leadingEquals, namesTaken } from './taken.js'

/** A cell of a book with `styles` styles. */
export function checkCell(v: unknown, path: string, styles: number): CellInput {
  const c = object(v, path, 'a cell object such as { value: 1 } or { formula: "A1*2" }')
  onlyKeys(c, path, CELL_FIELDS)
  const value = c.value
  // NaN and infinities have no place in a workbook, as for held and cached values
  if (typeof value === 'number' && !Number.isFinite(value)) throw new InputError(`${path}.value`, `expected a finite number, got ${what(value)}`)
  if (value !== undefined && value !== null && typeof value !== 'number' && typeof value !== 'string' && typeof value !== 'boolean') {
    if (!isObject(value) || Object.keys(value).length !== 1 || typeof value.date !== 'string')
      throw new InputError(`${path}.value`, `expected a number, string, boolean, null or { date: "YYYY-MM-DD" }, got ${what(value)}`)
    try {
      dateToSerial(value.date)
    } catch {
      throw new InputError(`${path}.value.date`, `invalid date ${JSON.stringify(value.date)}: expected an ISO date such as "2025-01-31" or "2025-01-31T18:00:00"`)
    }
  }
  if (c.meta !== undefined) checkMeta(c.meta, `${path}.meta`)
  if (c.formula !== undefined && typeof c.formula !== 'string') checkString(c.formula, `${path}.formula`)
  if (c.numberFormat !== undefined && typeof c.numberFormat !== 'string') checkFormat(c.numberFormat, `${path}.numberFormat`)
  if (c.style !== undefined) checkStyleIndex(c.style, `${path}.style`, styles)
  if (c.array !== undefined) checkString(c.array, `${path}.array`)
  if (c.note !== undefined) checkNote(c.note, `${path}.note`)
  if (c.link !== undefined) checkLink(c.link, `${path}.link`)
  if (c.held === null) throw new InputError(`${path}.held`, 'expected a number, text, a boolean or { error: code }, got null')
  if (c.held !== undefined) checkCellValue(c.held, `${path}.held`)
  return c as CellInput
}

/** A cell's note: its text and author. */
function checkNote(v: unknown, path: string) {
  const n = object(v, path, 'a note such as { author: "Ann", text: "Checked" }')
  onlyKeys(n, path, ['author', 'text'])
  checkString(n.text, `${path}.text`)
  if (n.author !== undefined) checkString(n.author, `${path}.author`)
}

/** A cell's hyperlink: a url, a location or both, and a tooltip. */
function checkLink(v: unknown, path: string) {
  const l = object(v, path, 'a hyperlink such as { url: "https://example.com/" } or { location: "Sheet2!B12" }')
  onlyKeys(l, path, ['url', 'location', 'tooltip'])
  for (const k of ['url', 'location', 'tooltip']) if (l[k] !== undefined) checkString(l[k], `${path}.${k}`)
  if (!l.url && !l.location) throw new InputError(path, 'a hyperlink needs a url, a location or both')
}

/** Names, their formulas without a leading `=`. */
export function checkNames(v: unknown, path: string): NameInput[] {
  const names = array(v, path, 'an array of names')
  names.forEach((n, i) => {
    const p = `${path}[${i}]`
    const o = object(n, p, "a name such as { name: 'rate', ref: \"'Data'!$B$1\" }")
    onlyKeys(o, p, ['name', 'ref', 'scope', 'hidden'])
    checkString(o.name, `${p}.name`)
    checkString(o.ref, `${p}.ref`)
    if (o.scope !== undefined) checkString(o.scope, `${p}.scope`)
    if (o.hidden !== undefined) checkBoolean(o.hidden, `${p}.hidden`)
  })
  return namesTaken(names as NameInput[])
}

/** Within the grid for sure (up to ZZ999999); anything else takes the full check. */
const SURE_ADDRESS = /^[A-Z]{1,2}[1-9][0-9]{0,5}$/

/**
 * A cell's meta, which the input hash writes as JSON and a thread takes by structured clone: a BigInt or a cycle breaks
 * both, so it is refused; what JSON turns into null or `{}` (a function, NaN, a Date) passes, as it did.
 */
function checkMeta(v: unknown, path: string, open: object[] = []) {
  if (typeof v === 'bigint') throw new InputError(path, `expected JSON, got a BigInt`)
  if (v === null || typeof v !== 'object') return
  if (open.includes(v)) throw new InputError(path, 'expected JSON, got a cycle: an object inside itself')
  open.push(v)
  for (const k of Object.keys(v)) checkMeta((v as Record<string, unknown>)[k], `${path}${Array.isArray(v) ? `[${k}]` : `.${k}`}`, open)
  open.pop()
}

/** The fields of a cell (`CellInput`). */
export const CELL_FIELDS = ['value', 'formula', 'numberFormat', 'style', 'meta', 'array', 'note', 'link', 'held']

/** The common cell, checked without building paths (the cell loop is hot); anything else takes the full check. */
function plainCell(c: unknown, styles: number): boolean {
  if (!isObject(c)) return false
  for (const k in c) {
    const v = c[k]
    // meta that is an object takes the full check, for BigInts and cycles
    if (k === 'meta' && (v === null || (typeof v !== 'object' && typeof v !== 'bigint'))) continue
    if (k === 'value' && (v === null || v === undefined || Number.isFinite(v) || typeof v === 'string' || typeof v === 'boolean')) continue
    // a formula with a leading `=` takes the full check, which marks it
    if (k === 'formula' && (v === undefined || (typeof v === 'string' && !leadingEquals(v)))) continue
    if (k === 'numberFormat' && (v === undefined || typeof v === 'string')) continue
    if (k === 'style' && Number.isInteger(v) && (v as number) >= 0 && (v as number) < styles) continue
    return false
  }
  return true
}

/** Why Excel would not take a sheet name, or null: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` (reserved, in any case). */
export function sheetNameProblem(name: string): Message | null {
  if (name === '') return { key: 'sheetNameEmpty' }
  if (name.length > 31) return { key: 'sheetNameLong', name, length: name.length }
  const bad = /[:\\/?*[\]]/.exec(name)
  if (bad) return { key: 'sheetNameChar', char: bad[0] }
  if (name.startsWith("'") || name.endsWith("'")) return { key: 'sheetNameApostrophe' }
  // Excel refuses to rename a sheet so (the struct set `history`)
  if (name.toLowerCase() === 'history') return { key: 'sheetNameHistory' }
  return null
}

const DATE_SYSTEMS = [1900, 1904]

const LINE = { c: { size: 'width', max: 255, what: 'columns', key: SPAN_KEY.c }, r: { size: 'height', max: 409.5, what: 'rows', key: SPAN_KEY.r } } as const

/** A column or row entry, or a patch of one (fields may be null). */
function checkLine(v: unknown, p: string, axis: Axis, styles: number, patch: boolean) {
  const { size, max } = LINE[axis]
  const o = object(v, p, `an object such as { ${size}: 20 }`)
  onlyKeys(o, p, [size, 'bestFit', 'style', 'numberFormat', 'hidden', 'outlineLevel', 'collapsed'])
  const field = (k: string, check: (x: unknown, path: string) => void) => givenIn(patch)(o, k) && check(o[k], `${p}.${k}`)
  field(size, (x, path) => checkNumber(x, path, 0, max))
  field('style', (x, path) => checkStyleIndex(x, path, styles))
  field('numberFormat', checkFormat)
  field('hidden', checkBoolean)
  field('outlineLevel', (x, path) => checkNumber(x, path, 1, 7, true))
  field('collapsed', checkBoolean)
  field('bestFit', checkBoolean)
}

function checkLines(v: unknown, path: string, axis: Axis, styles: number) {
  const map = object(v, path, `${LINE[axis].what} by ${LINE[axis].key}`)
  const spans = Object.entries(map).map(([key, e]): [number, number, string] => {
    const p = `${path}[${JSON.stringify(key)}]`
    checkLine(e, p, axis, styles, false)
    return [...checkSpan(key, p, axis), key]
  })
  spans.sort((a, b) => a[0] - b[0])
  for (let i = 1; i < spans.length; i++) if (spans[i][0] <= spans[i - 1][1]) throw new InputError(path, `${JSON.stringify(spans[i][2])} overlaps ${JSON.stringify(spans[i - 1][2])}`)
}

function checkMerges(v: unknown, path: string) {
  const list = array(v, path, 'an array of merged areas such as ["B2:D3"]')
  const xs = list.map((m, i) => {
    const x = typeof m === 'string' ? parseMerge(m) : null
    if (!x) throw new InputError(`${path}[${i}]`, `expected a merged area such as "B2:D3" (top-left to bottom-right, two cells or more), got ${what(m)}`)
    return x
  })
  const o = overlapping(xs)
  if (o) throw new InputError(`${path}[${o[1]}]`, `${list[o[1]]} overlaps ${list[o[0]]}`)
}

/** Frozen panes: rows, columns or both, each from 1 to one less than the sheet has. */
function checkFrozen(v: unknown, path: string) {
  const f = object(v, path, 'frozen panes such as { rows: 1 } or { rows: 2, columns: 1 }')
  onlyKeys(f, path, ['rows', 'columns'])
  if (f.rows === undefined && f.columns === undefined) throw new InputError(path, 'expected rows, columns or both')
  if (f.rows !== undefined) checkNumber(f.rows, `${path}.rows`, 1, MAX_ROWS - 1, true)
  if (f.columns !== undefined) checkNumber(f.columns, `${path}.columns`, 1, MAX_COLS - 1, true)
}

/** A sheet's AutoFilter: `{ range }`, a cell or a range from its top-left cell to its bottom-right one. */
function checkAutoFilter(v: unknown, path: string) {
  const f = object(v, path, 'an AutoFilter such as { range: "A1:D20" }')
  onlyKeys(f, path, ['range'])
  const x = typeof f.range === 'string' ? parseRange(f.range) : null
  if (!x || (f.range !== rectText(x) && f.range !== cornersText(x)))
    throw new InputError(`${path}.range`, `expected a cell or a range from its top-left cell to its bottom-right one such as "A1:D20", got ${what(f.range)}`)
}

/** Iterative calculation's bounds (`BookInput.iteration`): 1 to 32,767 rounds, as Excel takes them, a change from 0. */
export const ITERATION = { maxIterations: [1, 32_767], maxChange: [0, Number.MAX_VALUE] } as const

/** A hidden sheet's states (`SheetInput.hidden`) with the names an xlsx file gives them (`<sheet state>`). */
export const SHEET_STATES = new Map<true | 'very', string>([
  [true, 'hidden'],
  ['very', 'veryHidden'],
])

/**
 * The book: sheets with unique names by Excel's rules (compared case-insensitively), one of them visible at least, cells
 * by A1 address, their looks, columns, rows, merged areas, AutoFilters, frozen panes, data validation and outline settings, names,
 * styles. It is returned as the engine takes it: its formulas without a leading `=`, its number formats by the name
 * `numberFormat` (a copy of what needs it, `taken.ts`), the entries of a look alone in its sheets' looks (`looksTaken`);
 * with `looked`, those of the sheets that have looks only (a writer keeps a sheet without looks as it is given). `keys`
 * gets the addresses of each sheet's cells, in their order, where the book returned keeps the cells object given.
 */
export function checkBook(v: unknown, path = 'book', looked = false, keys?: Map<object, string[]>): BookInput {
  const book = object(v, path, 'a workbook such as { sheets: [{ name: "Sheet1", cells: {} }] }')
  onlyKeys(book, path, ['sheets', 'names', 'styles', 'locale', 'dateSystem', 'iteration', 'externalBooks', 'activeSheet', 'font'])
  if (book.locale !== undefined) checkBookLocale(book.locale, `${path}.locale`)
  if (book.font !== undefined) {
    const f = object(book.font, `${path}.font`, "a font such as { name: 'Arial', size: 10 }")
    onlyKeys(f, `${path}.font`, ['name', 'size'])
    if (checkString(f.name, `${path}.font.name`) === '') throw new InputError(`${path}.font.name`, 'expected the name of a font, got an empty string')
    checkNumber(f.size, `${path}.font.size`, 1, 409)
  }
  if (book.dateSystem !== undefined) checkOneOf(book.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (book.iteration !== undefined) {
    const it = object(book.iteration, `${path}.iteration`, 'iteration settings such as { maxIterations: 100, maxChange: 0.001 }')
    onlyKeys(it, `${path}.iteration`, ['maxIterations', 'maxChange'])
    if (it.maxIterations !== undefined) checkNumber(it.maxIterations, `${path}.iteration.maxIterations`, ...ITERATION.maxIterations, true)
    if (it.maxChange !== undefined) checkNumber(it.maxChange, `${path}.iteration.maxChange`, ...ITERATION.maxChange)
  }
  const styles = book.styles === undefined ? 0 : checkStyles(book.styles, `${path}.styles`).length
  const seen = new Map<string, { name: string; i: number }>()
  // the cells of each sheet that the engine takes otherwise: a formula with a leading `=`, a look alone (into the sheet's looks)
  const taken: string[][] = []
  const alone: string[][] = []
  const listed: string[][] = []
  array(book.sheets, `${path}.sheets`, 'an array of sheets').forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(s, p, 'a sheet such as { name: "Sheet1", cells: {} }')
    onlyKeys(sheet, p, ['name', 'cells', 'looks', 'columns', 'rows', 'merges', 'filtered', 'autoFilter', 'tables', 'frozen', 'conditionalFormats', 'validations', 'outline', 'hidden', 'view', 'protection', 'tabColor'])
    const name = checkString(sheet.name, `${p}.name`)
    const problem = sheetNameProblem(name)
    if (problem) throw new InputError(`${p}.name`, english(problem))
    const dup = seen.get(sheetKey(name))
    if (dup) throw new InputError(`${p}.name`, `duplicate sheet name ${JSON.stringify(name)}: ${path}.sheets[${dup.i}] is ${JSON.stringify(dup.name)} (sheet names are case-insensitive)`)
    seen.set(sheetKey(name), { name, i })
    const cells = object(sheet.cells, `${p}.cells`, 'an object of cells by address such as { A1: { value: 1 } }')
    // the cells of old-kind array formulas, checked against the sheet once all its cells are
    const arrays: string[] = []
    for (const addr of (listed[i] = Object.keys(cells))) {
      const cell = cells[addr]
      if (!(SURE_ADDRESS.test(addr) && plainCell(cell, styles))) {
        const cp = `${p}.cells[${JSON.stringify(addr)}]`
        checkAddress(addr, cp)
        const c = checkCell(cell, cp, styles)
        if (c.array !== undefined) arrays.push(addr)
        if (leadingEquals(c.formula)) (taken[i] ??= []).push(addr)
      }
      if (lookAlone(cell as CellInput)) (alone[i] ??= []).push(addr)
    }
    if (arrays.length) checkArrays(cells as Record<string, CellInput>, arrays, p)
    if (sheet.looks !== undefined) checkLooks(sheet.looks, `${p}.looks`, styles)
    if (sheet.columns !== undefined) checkLines(sheet.columns, `${p}.columns`, 'c', styles)
    if (sheet.rows !== undefined) checkLines(sheet.rows, `${p}.rows`, 'r', styles)
    if (sheet.merges !== undefined) checkMerges(sheet.merges, `${p}.merges`)
    if (sheet.filtered !== undefined) checkBoolean(sheet.filtered, `${p}.filtered`)
    if (sheet.autoFilter !== undefined) checkAutoFilter(sheet.autoFilter, `${p}.autoFilter`)
    if (sheet.view !== undefined) checkSheetView(sheet.view, `${p}.view`)
    if (sheet.tabColor !== undefined) checkColor(sheet.tabColor, `${p}.tabColor`)
    if (sheet.frozen !== undefined) checkFrozen(sheet.frozen, `${p}.frozen`)
    if (sheet.conditionalFormats !== undefined) checkConditionalFormats(sheet.conditionalFormats, `${p}.conditionalFormats`)
    if (sheet.validations !== undefined) checkValidations(sheet.validations, `${p}.validations`)
    if (sheet.outline !== undefined) checkOutline(sheet.outline, `${p}.outline`)
    if (sheet.protection !== undefined) checkSheetProtection(sheet.protection, `${p}.protection`)
    if (sheet.hidden !== undefined && sheet.hidden !== false && !SHEET_STATES.has(sheet.hidden as true)) throw new InputError(`${p}.hidden`, `expected true, false or "very", got ${what(sheet.hidden)}`)
  })
  // Excel keeps one visible sheet at least
  const sheets = book.sheets as { hidden?: unknown }[]
  if (sheets.length && sheets.every((s) => s.hidden !== undefined && s.hidden !== false)) throw new InputError(`${path}.sheets`, 'a workbook must contain at least one visible sheet: every sheet is hidden')
  if (book.activeSheet !== undefined) {
    const name = checkString(book.activeSheet, `${path}.activeSheet`)
    const active = (book.sheets as BookInput['sheets']).find((s) => sheetKey(s.name) === sheetKey(name))
    if (!active || active.hidden) throw new InputError(`${path}.activeSheet`, 'expected the name of a visible sheet of the workbook')
  }
  if (book.names !== undefined) checkNames(book.names, `${path}.names`)
  if (book.externalBooks !== undefined) checkExternalBooks(book.externalBooks, `${path}.externalBooks`)
  checkTables(book as unknown as BookInput, path)
  const out = bookTaken(book as unknown as BookInput, taken)
  const sheetsTaken = out.sheets.map((s, i) => (looked && !hasLooks(s) ? s : looksTaken(s, alone[i] ?? [])))
  sheetsTaken.forEach((s, i) => s.cells === (book.sheets as BookInput['sheets'])[i].cells && keys?.set(s.cells, listed[i]))
  return sheetsTaken.every((s, i) => s === out.sheets[i]) ? out : { ...out, sheets: sheetsTaken }
}

/** A sheet's looks: rectangles in A1 style (`B2:D9`, `B:D`, `2:9`, `B2`), each with a style of the book and a number format, both optional. */
function checkLooks(v: unknown, path: string, styles: number) {
  array(v, path, 'an array of formatting ranges such as [{ range: "A2:XFD9", style: 0 }]').forEach((x, k) => {
    const p = `${path}[${k}]`
    const o = object(x, p, 'a formatting range such as { range: "A2:XFD9", style: 0, numberFormat: "0.00" }')
    onlyKeys(o, p, ['range', 'style', 'numberFormat'])
    checkRange(o.range, `${p}.range`, false)
    if (o.style !== undefined) checkStyleIndex(o.style, `${p}.style`, styles)
    if (o.numberFormat !== undefined) checkFormat(o.numberFormat, `${p}.numberFormat`)
  })
}

/**
 * Other workbooks a book links to (`BookInput.externalBooks`): each a target, its sheets by name with their stored values
 * by address (a number, text, a boolean, an error, null for a cell stored empty) and the mark of a failed update, its
 * names with what they refer to and the index of the sheet of a sheet's name. Without `values` the stored values are
 * not read, only that a sheet's `cells` is an object: the options of the calls that show, read or check formula text
 * use the links' targets, sheets and names alone (the package's own calls over a checked book skip the check).
 */
export function checkExternalBooks(v: unknown, path: string, values = true): ExternalBook[] {
  return array(v, path, 'an array of linked workbooks such as { target: "Book2.xlsx", sheets: [{ name: "Sheet1", cells: { A1: 5 } }] }').map((b, i) => {
    const p = `${path}[${i}]`
    const book = object(b, p, 'a linked workbook such as { target: "Book2.xlsx", sheets: [{ name: "Sheet1", cells: { A1: 5 } }] }')
    onlyKeys(book, p, ['target', 'sheets', 'names'])
    checkString(book.target, `${p}.target`)
    const sheets = array(book.sheets, `${p}.sheets`, 'an array of sheets such as { name: "Sheet1", cells: { A1: 5 } }')
    sheets.forEach((s, j) => {
      const sp = `${p}.sheets[${j}]`
      const sheet = object(s, sp, 'a sheet such as { name: "Sheet1", cells: { A1: 5 } }')
      onlyKeys(sheet, sp, ['name', 'cells', 'refreshError'])
      checkString(sheet.name, `${sp}.name`)
      if (sheet.refreshError !== undefined && sheet.refreshError !== true) throw new InputError(`${sp}.refreshError`, `expected true, got ${what(sheet.refreshError)}`)
      if (sheet.cells === undefined) return
      const cells = object(sheet.cells, `${sp}.cells`, 'an object of stored values by address such as { A1: 5 }')
      if (values)
        for (const addr of Object.keys(cells)) {
          const cp = `${sp}.cells[${JSON.stringify(addr)}]`
          if (!SURE_ADDRESS.test(addr)) checkAddress(addr, cp)
          checkCellValue(cells[addr], cp)
        }
    })
    if (book.names !== undefined)
      array(book.names, `${p}.names`, 'an array of names such as { name: "Rate", ref: "Sheet1!$B$1" }').forEach((n, j) => {
        const np = `${p}.names[${j}]`
        const name = object(n, np, 'a name such as { name: "Rate", ref: "Sheet1!$B$1" }')
        onlyKeys(name, np, ['name', 'ref', 'sheet'])
        checkString(name.name, `${np}.name`)
        checkString(name.ref, `${np}.ref`)
        if (name.sheet !== undefined) checkNumber(name.sheet, `${np}.sheet`, 0, sheets.length - 1, true)
      })
    return book as unknown as ExternalBook
  })
}

/** ISO 8601 with Z or an offset, or a date alone: the forms `Date.parse` reads the same in every time zone. */
const MOMENT = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(?:Z|([+-])(\d{2}):(\d{2})))?$/

/** A moment the same in every time zone, on a day the calendar has (`Date.parse` rolls 30 February over to March). */
function checkMoment(v: unknown, path: string): string {
  const now = checkString(v, path)
  const m = MOMENT.exec(now)
  const ms = Date.parse(now)
  // the moment back in the given offset must show the given date and time
  const shown = m && !Number.isNaN(ms) ? new Date(ms + (m[7] === undefined ? 0 : (m[7] === '-' ? -1 : 1) * (Number(m[8]) * 60 + Number(m[9])) * 60_000)) : undefined
  const parts = shown && [shown.getUTCFullYear(), shown.getUTCMonth() + 1, shown.getUTCDate(), shown.getUTCHours(), shown.getUTCMinutes(), shown.getUTCSeconds()]
  if (!parts || parts.some((p, i) => m![i + 1] !== undefined && p !== Number(m![i + 1])))
    throw new InputError(path, `invalid moment ${JSON.stringify(now)}: expected ISO 8601 with Z or an offset, such as "2026-01-31T12:00:00Z", or a date such as "2026-01-31"`)
  return now
}

export function checkClock(v: unknown, path: string): Clock {
  const c = object(v, path, 'a clock such as { now: "2026-01-31T12:00:00Z", seed: 1 }')
  onlyKeys(c, path, ['now', 'seed'])
  checkMoment(c.now, `${path}.now`)
  if (typeof c.seed !== 'number' || !Number.isFinite(c.seed)) throw new InputError(`${path}.seed`, `expected a finite number, got ${what(c.seed)}`)
  return c as unknown as Clock
}

function checkExternals(v: unknown, path: string): Externals {
  const e = object(v, path, 'externals such as { functions: ["BALANCE"], values: {} }')
  onlyKeys(e, path, ['functions', 'values'])
  array(e.functions, `${path}.functions`, 'an array of function names').forEach((f, i) => checkString(f, `${path}.functions[${i}]`))
  if (e.values !== undefined) {
    const values = object(e.values, `${path}.values`, 'an object of values by call text')
    // a non-finite number stays the #NUM! it computes to
    for (const k in values) if (typeof values[k] !== 'number') checkCellValue(values[k], `${path}.values[${JSON.stringify(k)}]`)
  }
  return e as unknown as Externals
}

const EXECUTIONS: readonly Execution[] = ['closures', 'eval']

function checkExecution(v: unknown, path: string) {
  if (!EXECUTIONS.includes(v as Execution)) throw new InputError(path, `expected "closures" or "eval", got ${what(v)}`)
}

const CLOCK_HINT = ' (pass the clock as { clock: { now, seed } })'

const LIMITS = ['formulas', 'cells', 'spill', 'writes']

/** `limits`: whole numbers of at least 1. */
export function checkLimits(v: unknown, path: string): CalcLimits {
  const o = object(v, path, 'limits such as { formulas: 100000, cells: 10000000 }')
  onlyKeys(o, path, LIMITS)
  for (const k of LIMITS) if (o[k] !== undefined) checkNumber(o[k], `${path}.${k}`, 1, Number.MAX_SAFE_INTEGER, true)
  return o as CalcLimits
}

/** `cached`: a file's stored values in the shape of a result, each a cell value. */
function checkCached(v: unknown, path: string): CachedValues {
  const o = object(v, path, "stored values such as readXlsx's cached, { sheets: [{ name, cells: { B1: { value: 2 } } }] }")
  array(o.sheets, `${path}.sheets`, 'an array of sheets such as { name: "Sheet1", cells: { B1: { value: 2 } } }').forEach((x, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(x, p, 'a sheet such as { name: "Sheet1", cells: { B1: { value: 2 } } }')
    checkString(sheet.name, `${p}.name`)
    const cells = object(sheet.cells, `${p}.cells`, 'stored values by address such as { B1: { value: 2 } }')
    for (const addr in cells) {
      const cp = `${p}.cells[${JSON.stringify(addr)}]`
      checkCellValue(object(cells[addr], cp, 'a stored value such as { value: 2 }').value, `${cp}.value`)
    }
  })
  return o as unknown as CachedValues
}

/** The run fields the compute and session options both check, each with its checker. */
const RUN_FIELDS: Record<string, (v: unknown, path: string) => void> = { cached: checkCached, clock: checkClock, limits: checkLimits, externals: checkExternals, execution: checkExecution, display: checkBoolean }

/** The run fields `keys`, each checked when given, in the caller's own order (undo and the live clock stay with their calls). */
const runFields = (o: Record<string, unknown>, path: string, keys: readonly string[]) => {
  for (const k of keys) if (o[k] !== undefined) RUN_FIELDS[k](o[k], `${path}.${k}`)
}

export function checkComputeOptions(v: unknown, path = 'opts'): ComputeOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } }')
  onlyOptions(o, path, ['clock', 'externals', 'execution', 'display', 'limits', 'cached'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  runFields(o, path, ['cached', 'clock', 'limits', 'externals', 'execution', 'display'])
  return o as ComputeOptions
}

export function checkSessionOptions(v: unknown, path = 'opts'): SessionOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } } or { live: true }')
  onlyOptions(o, path, ['clock', 'live', 'externals', 'execution', 'display', 'limits', 'undo', 'cached'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.live !== undefined) checkBoolean(o.live, `${path}.live`)
  runFields(o, path, ['cached', 'limits'])
  if (o.undo !== undefined) {
    const u = object(o.undo, `${path}.undo`, 'undo options such as { depth: 100, maxMb: 64 }')
    onlyKeys(u, `${path}.undo`, ['depth', 'maxMb'])
    if (u.depth !== undefined) checkNumber(u.depth, `${path}.undo.depth`, 0, 1_000_000, true)
    if (u.maxMb !== undefined) checkNumber(u.maxMb, `${path}.undo.maxMb`, 0, 1_000_000)
  }
  runFields(o, path, ['display'])
  if (o.live === true && o.clock !== undefined) throw new InputError(path, 'a live session takes its clock from the system: pass either live or clock')
  runFields(o, path, ['clock', 'externals', 'execution'])
  return o as SessionOptions
}

/**
 * Each of a list of edits checked as an edit (`{ sheet, cell, input }`): a sheet of the book (`sheetIndex` resolves it),
 * a valid address and an input, which `check` checks unless null; `missing` says what an edit without one lacks.
 */
export function checkEditList(v: unknown, sheetIndex: (name: string) => number | undefined, path: string, missing: string, check: (input: unknown, path: string) => void): CellEdit[] {
  return array(v, path, 'an array of edits such as [{ sheet: "Sheet1", cell: "A1", input: { value: 1 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'an edit such as { sheet: "Sheet1", cell: "A1", input: { value: 1 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'input'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    if (o.input === undefined) throw new InputError(`${p}.input`, `expected ${missing}`)
    if (o.input !== null) check(o.input, `${p}.input`)
    return o as unknown as CellEdit
  })
}

/** Edits of cells of a book with `styles` styles, as the engine takes them (`taken.ts`). */
export function checkEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): CellEdit[] {
  return editsTaken(checkEditList(v, sheetIndex, path, 'a cell object, or null to clear the cell', (input, p) => checkCell(input, p, styles)))
}

/** The fields of a sheet a layout edit sets, in the order an edit is read by, with their checks; those that may be null, which removes them. */
const SHEET_EDITS: Record<string, { check: (v: unknown, path: string) => void; nullable?: true }> = {
  merges: { check: checkMerges },
  filtered: { check: checkBoolean },
  autoFilter: { check: checkAutoFilter, nullable: true },
  protection: { check: checkSheetProtection, nullable: true },
  frozen: { check: checkFrozen, nullable: true },
  tabColor: { check: checkColor, nullable: true },
  conditionalFormats: { check: checkConditionalFormats },
  validations: { check: checkValidations },
  outline: { check: checkOutline, nullable: true },
}
const EDITED = Object.keys(SHEET_EDITS)

/**
 * Layout edits of a book with `styles` styles: a sheet of the book, and columns, rows, merged areas, the filter mode, an
 * AutoFilter, frozen panes, rules of data validation, outline settings or protection; as the engine takes them (`taken.ts`).
 */
export function checkLayoutEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): LayoutEdit[] {
  const edits = array(v, path, 'an array of layout edits such as [{ sheet: "Sheet1", columns: "C", set: { width: 20 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'a layout edit such as { sheet: "Sheet1", columns: "C", set: { width: 20 } }')
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    // an edit of one field of the sheet, checked by the field's check; null removes a field that may go
    const field = EDITED.find((k) => o[k] !== undefined || (SHEET_EDITS[k].nullable && k in o))
    if (field) {
      onlyKeys(o, p, ['sheet', field])
      if (o[field] !== null || !SHEET_EDITS[field].nullable) SHEET_EDITS[field].check(o[field], `${p}.${field}`)
      return o as unknown as LayoutEdit
    }
    if (o.fit !== undefined) {
      const lines = o.rows !== undefined ? 'rows' : 'columns'
      onlyKeys(o, p, ['sheet', lines, 'fit'])
      if (o.fit !== true) throw new InputError(`${p}.fit`, `expected true, got ${what(o.fit)}`)
      checkSpan(o[lines], `${p}.${lines}`, lines === 'rows' ? 'r' : 'c')
      return o as unknown as LayoutEdit
    }
    const axis: Axis = o.rows !== undefined ? 'r' : 'c'
    const lines = LINES[axis]
    onlyKeys(o, p, ['sheet', lines, 'set'], o.columns === undefined && o.rows === undefined ? ` (one of columns, rows, ${EDITED.slice(0, -1).join(', ')} or ${EDITED.at(-1)})` : '')
    checkSpan(o[lines], `${p}.${lines}`, axis)
    checkLine(o.set, `${p}.set`, axis, styles, true)
    return o as unknown as LayoutEdit
  })
  return layoutTaken(edits)
}

/** A sheet name known to the book; returns its index. */
export function checkSheet(v: unknown, sheetIndex: (name: string) => number | undefined, path: string): number {
  const i = sheetIndex(checkString(v, path))
  if (i === undefined) throw new InputError(path, `unknown sheet ${JSON.stringify(v)}`)
  return i
}

/** A sheet of a book, by its name in any case, as the book spells it; `InputError` naming `path` when it has none. */
export function checkSheetOf<S extends { name: string }>(book: { sheets: S[] }, v: unknown, path: string): S {
  const name = checkString(v, path)
  const sheet = book.sheets.find((s) => sheetKey(s.name) === sheetKey(name))
  if (!sheet) throw new InputError(path, `unknown sheet ${JSON.stringify(v)}`)
  return sheet
}

export function checkValidateContext(v: unknown, path = 'ctx'): ValidateContext | undefined {
  if (v === undefined) return undefined
  const o = object(v, path, 'a context such as { sheets: ["Sheet1"], names: [] }')
  onlyKeys(o, path, ['sheets', 'names', 'externalFunctions', 'tables', 'table', 'externalBooks'])
  if (o.table !== undefined) checkString(o.table, `${path}.table`)
  if (o.externalBooks !== undefined) checkExternalBooks(o.externalBooks, `${path}.externalBooks`, false)
  for (const k of ['sheets', 'names', 'externalFunctions'] as const)
    if (o[k] !== undefined) array(o[k], `${path}.${k}`, 'an array of strings').forEach((s, i) => checkString(s, `${path}.${k}[${i}]`))
  if (o.tables !== undefined)
    array(o.tables, `${path}.tables`, 'an array of tables such as { name: "Sales", columns: ["Qty"] }').forEach((t, i) => {
      const p = `${path}.tables[${i}]`
      const x = object(t, p, 'a table such as { name: "Sales", columns: ["Qty"] }')
      onlyKeys(x, p, ['name', 'columns'])
      checkString(x.name, `${p}.name`)
      array(x.columns, `${p}.columns`, 'an array of column names').forEach((c, j) => checkString(c, `${p}.columns[${j}]`))
    })
  return o as ValidateContext
}

/** `validateName`'s options: the book's locale. */
export function checkNameOptions(v: unknown, path = 'opts'): NameOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { locale: 'ru-RU' }")
  onlyKeys(o, path, ['locale'])
  optionalLocale(o.locale, `${path}.locale`)
  return o as NameOptions
}

/** `validateNames`'s context: the book's sheets, which a name's scope needs, its tables, external functions and locale. */
export function checkNamesContext(v: unknown, path = 'ctx'): NamesContext {
  const o = object(v, path, "a context such as { sheets: ['Data'], locale: 'ru-RU' }")
  onlyKeys(o, path, ['sheets', 'tables', 'externalFunctions', 'locale', 'externalBooks'])
  if (o.sheets === undefined) throw new InputError(`${path}.sheets`, "expected the workbook's sheet names as an array of strings")
  const { locale, ...rest } = o
  checkValidateContext(rest, path)
  optionalLocale(locale, `${path}.locale`)
  return o as unknown as NamesContext
}

/** `displayValue`'s value: a number, text, a boolean, null or an error `{ error: code }`. */
export function checkCellValue(v: unknown, path: string): CellValue {
  if (v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) return v
  if (typeof v === 'object' && !Array.isArray(v) && Object.keys(v).length === 1 && Object.hasOwn(ERROR_TEXT, (v as { error?: unknown }).error as string)) return v as CellValue
  throw new InputError(path, `expected a number, text, a boolean, null or { error: code }, got ${what(v)}`)
}

/** `displayValue`'s options. */
export function checkDisplayOptions(v: unknown, path = 'opts'): DisplayOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { numberFormat: '#,##0.00', locale: 'ru-RU', width: 12 }")
  onlyKeys(o, path, ['numberFormat', 'locale', 'width', 'dateSystem'])
  if (o.numberFormat !== undefined) checkFormat(o.numberFormat, `${path}.numberFormat`)
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  optionalLocale(o.locale, `${path}.locale`)
  if (o.width !== undefined) checkNumber(o.width, `${path}.width`, 0, 255 * MAX_COLS)
  return o as DisplayOptions
}

/** `entryWidth`'s options: `displayValue`'s and the column's `bestFit`. */
export function checkEntryWidthOptions(v: unknown, path = 'opts'): EntryWidthOptions {
  if (v === undefined) return {}
  const { bestFit, ...rest } = object(v, path, "options such as { numberFormat: 14, locale: 'ru-RU', width: 9.5, bestFit: true }")
  if (bestFit !== undefined) checkBoolean(bestFit, `${path}.bestFit`)
  return { ...checkDisplayOptions(rest, path), ...(bestFit !== undefined ? { bestFit: bestFit as boolean } : {}) }
}

/** The locale, date system and moment of `entryText`'s options (unchecked there), as entry options `checkEntryOptions` checks. */
export const entryOptionsOf = (o: EntryTextOptions): EntryOptions => ({ ...(o.locale !== undefined ? { locale: o.locale } : {}), ...(o.dateSystem !== undefined ? { dateSystem: o.dateSystem } : {}), ...(o.now !== undefined ? { now: o.now } : {}) })

/** `entryText`'s and `needsApostrophe`'s options: the cell's number format, the book's locale and date system, the moment for a date without a year. */
export function checkEntryTextOptions(v: unknown, path = 'opts'): EntryTextOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { numberFormat: '0%', locale: 'ru-RU' }")
  onlyKeys(o, path, ['numberFormat', 'locale', 'dateSystem', 'now'])
  if (o.numberFormat !== undefined) checkFormat(o.numberFormat, `${path}.numberFormat`)
  checkEntryOptions(entryOptionsOf(o as EntryTextOptions), path)
  return o as EntryTextOptions
}

/**
 * `parseEntry`'s options; of the cell only its number format, style, meta, note and link are read (it may be a
 * `CellView`), and the old name of the number format is refused.
 */
export function checkEntryOptions(v: unknown, path = 'opts'): EntryOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { locale: 'ru-RU', cell: { numberFormat: '0.00' } }")
  onlyKeys(o, path, ['locale', 'dateSystem', 'cell', 'now', 'externalBooks', 'inPlace'])
  optionalLocale(o.locale, `${path}.locale`)
  if (o.inPlace !== undefined) checkBoolean(o.inPlace, `${path}.inPlace`)
  if (o.externalBooks !== undefined) checkExternalBooks(o.externalBooks, `${path}.externalBooks`, false)
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (o.now !== undefined) checkMoment(o.now, `${path}.now`)
  if (o.cell !== undefined && o.cell !== null) {
    const c = object(o.cell, `${path}.cell`, 'a cell such as { numberFormat: "0.00", style: 1 }')
    // a cell's other fields are not read (a `CellView` has more), but the old name of its number format is refused
    if (Object.hasOwn(c, 'format')) onlyKeys({ format: c.format }, `${path}.cell`, ['numberFormat', 'style', 'meta', 'note', 'link'])
    if (c.numberFormat !== undefined) checkFormat(c.numberFormat, `${path}.cell.numberFormat`)
    if (c.style !== undefined && !(Number.isInteger(c.style) && (c.style as number) >= 0)) throw new InputError(`${path}.cell.style`, `expected the index of a style, got ${what(c.style)}`)
    if (c.note !== undefined) checkNote(c.note, `${path}.cell.note`)
    if (c.link !== undefined) checkLink(c.link, `${path}.cell.link`)
  }
  return o as EntryOptions
}
