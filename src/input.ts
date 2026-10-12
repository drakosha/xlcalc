/**
 * Checks of caller input at the API boundary. Wrong input (a value of the wrong type, a malformed cell
 * address, duplicate sheets, an unknown sheet in an edit) throws `InputError` naming the place and the
 * problem. Formula problems are not input errors: an invalid formula gives an error value in its cell.
 */
import type { AlignmentStyle, BookInput, BorderLine, CalcLimits, CellEdit, CellInput, CellStyle, CellValue, Clock, ComputeOptions, DisplayOptions, EntryOptions, EntryTextOptions, EntryWidthOptions, Execution, ExternalBook, Externals, LayoutEdit, NameInput, NameOptions, NamesContext, SessionOptions, ValidateContext } from './contract.js'
import { overlapping, parseMerge } from './layout/merges.js'
import { parseSpan } from './layout/spans.js'
import { english } from './locale/english.js'
import type { Message } from './locale/messages.js'
import { checkBookLocale, checkLocaleArg } from './locale/load.js'
import { dateToSerial, MAX_COLS, MAX_ROWS } from './model.js'
import { isBuiltin } from './numfmt/builtin.js'
import type { Axis } from './structure/area.js'
import { checkArrays } from './array-areas.js'
import { ERROR_TEXT, sheetKey } from './syntax.js'
import { checkTables } from './tables/check.js'
import { InputError, what } from './input-error.js'
import { optionHint } from './strict-args.js'
import { checkConditionalFormats } from './cond/check.js'
import { checkOutline, checkValidations } from './validation/check.js'
import { bookTaken, editsTaken, layoutTaken, leadingEquals, namesTaken } from './taken.js'

export { InputError, what } from './input-error.js'

/** A plain object: not null, not an array. */
export const isObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v)

export function object(v: unknown, path: string, expected: string): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

export function array(v: unknown, path: string, expected: string): unknown[] {
  if (!Array.isArray(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

export function checkString(v: unknown, path: string): string {
  if (typeof v !== 'string') throw new InputError(path, `expected a string, got ${what(v)}`)
  return v
}

/** Fields by their old names, refused since they were renamed, and the name of each now. */
const RENAMED: Record<string, string> = { format: 'numberFormat' }

/** The fields of `o` are of `allowed`: an unknown one is refused, with the name now of a field by its old name (`format`). */
export function onlyKeys(o: Record<string, unknown>, path: string, allowed: readonly string[], hint = '') {
  for (const k of Object.keys(o))
    if (!allowed.includes(k)) {
      const now = Object.hasOwn(RENAMED, k) && allowed.includes(RENAMED[k]) ? ` (${k} is the old name of ${RENAMED[k]})` : ''
      throw new InputError(path, `unknown field ${JSON.stringify(k)} (expected ${allowed.join(', ')})${hint}${now}`)
    }
}

/** An options object's fields, as `onlyKeys`: an unknown one that another call takes says which call (`optionHint`). */
export const onlyOptions = (o: Record<string, unknown>, path: string, allowed: readonly string[], hint = '') => onlyKeys(o, path, allowed, hint + optionHint(o, allowed))

const ADDRESS = /^([A-Z]{1,3})([1-9][0-9]*)$/
const MAX_ADDRESS = 'XFD1048576'

/** A cell address in A1 style within the Excel grid. */
export function checkAddress(addr: unknown, path: string): string {
  if (typeof addr !== 'string') throw new InputError(path, `expected a cell address such as "B12", got ${what(addr)}`)
  const m = ADDRESS.exec(addr)
  if (!m) {
    const hint = ADDRESS.test(addr.toUpperCase()) ? ` (column letters must be upper-case: "${addr.toUpperCase()}")` : ''
    throw new InputError(path, `invalid cell address ${JSON.stringify(addr)}: expected A1 style such as "B12"${hint}`)
  }
  if (m[1].length < 3 && m[2].length < 7) return addr
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  if (col > MAX_COLS || Number(m[2]) > MAX_ROWS) throw new InputError(path, `cell address ${JSON.stringify(addr)} is outside the sheet (the last cell is ${MAX_ADDRESS})`)
  return addr
}

/** A number format: a code, or the number of a built-in format (0 to 49). */
export function checkFormat(v: unknown, path: string) {
  if (typeof v !== 'string' && !(typeof v === 'number' && isBuiltin(v))) throw new InputError(path, `expected a format code or the number of a built-in format, 0 to 49, got ${what(v)}`)
}

/** A style's fields where a cell takes its number: where such a style goes. */
const STYLE_HINT = " (a cell names a style of the book's styles by its index; a session's patchCells takes one by its fields)"

/** A style number of a book with `styles` styles. */
export function checkStyleIndex(v: unknown, path: string, styles: number) {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v >= styles) {
    const hint = isObject(v) ? STYLE_HINT : ''
    throw new InputError(path, (styles ? `expected the index of a style of the book, 0 to ${styles - 1}, got ${what(v)}` : `the book has no styles, got ${what(v)}`) + hint)
  }
}

/** A cell of a book with `styles` styles. */
export function checkCell(v: unknown, path: string, styles: number): CellInput {
  const c = object(v, path, 'a cell object such as { value: 1 } or { formula: "A1*2" }')
  onlyKeys(c, path, CELL_FIELDS)
  const value = c.value
  if (value !== undefined && value !== null && typeof value !== 'number' && typeof value !== 'string' && typeof value !== 'boolean') {
    if (!isObject(value) || Object.keys(value).length !== 1 || typeof value.date !== 'string')
      throw new InputError(`${path}.value`, `expected a number, string, boolean, null or { date: "YYYY-MM-DD" }, got ${what(value)}`)
    try {
      dateToSerial(value.date)
    } catch {
      throw new InputError(`${path}.value.date`, `invalid date ${JSON.stringify(value.date)}: expected ISO such as "2025-01-31" or "2025-01-31T18:00:00"`)
    }
  }
  if (c.formula !== undefined && typeof c.formula !== 'string') checkString(c.formula, `${path}.formula`)
  if (c.numberFormat !== undefined && typeof c.numberFormat !== 'string') checkFormat(c.numberFormat, `${path}.numberFormat`)
  if (c.style !== undefined) checkStyleIndex(c.style, `${path}.style`, styles)
  if (c.array !== undefined) checkString(c.array, `${path}.array`)
  if (c.note !== undefined) checkNote(c.note, `${path}.note`)
  if (c.link !== undefined) checkLink(c.link, `${path}.link`)
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
  if (!l.url && !l.location) throw new InputError(path, 'a hyperlink has a url, a location or both')
}

/** Names, their formulas without a leading `=`. */
export function checkNames(v: unknown, path: string): NameInput[] {
  const names = array(v, path, 'an array of names')
  names.forEach((n, i) => {
    const p = `${path}[${i}]`
    const o = object(n, p, "a name such as { name: 'rate', ref: \"'Data'!$B$1\" }")
    onlyKeys(o, p, ['name', 'ref', 'scope'])
    checkString(o.name, `${p}.name`)
    checkString(o.ref, `${p}.ref`)
    if (o.scope !== undefined) checkString(o.scope, `${p}.scope`)
  })
  return namesTaken(names as NameInput[])
}

/** Within the grid for sure (up to ZZ999999); anything else takes the full check. */
const SURE_ADDRESS = /^[A-Z]{1,2}[1-9][0-9]{0,5}$/

/** The fields of a cell (`CellInput`). */
export const CELL_FIELDS = ['value', 'formula', 'numberFormat', 'style', 'meta', 'array', 'note', 'link']

/** The common cell, checked without building paths (the cell loop is hot); anything else takes the full check. */
function plainCell(c: unknown, styles: number): boolean {
  if (!isObject(c)) return false
  for (const k in c) {
    const v = c[k]
    if (k === 'meta') continue
    if (k === 'value' && (v === null || v === undefined || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean')) continue
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

const COLOR = /^#[0-9A-Fa-f]{6}$/

export function checkColor(v: unknown, path: string) {
  if (typeof v !== 'string' || !COLOR.test(v)) throw new InputError(path, `expected a colour such as "#FF0000", got ${what(v)}`)
}

export function checkOneOf(v: unknown, path: string, values: readonly (string | number)[]) {
  if (!values.includes(v as string)) throw new InputError(path, `expected one of ${values.map((x) => JSON.stringify(x)).join(', ')}, got ${what(v)}`)
}

const DATE_SYSTEMS = [1900, 1904]

export function checkBoolean(v: unknown, path: string) {
  if (typeof v !== 'boolean') throw new InputError(path, `expected a boolean, got ${what(v)}`)
}

export function checkNumber(v: unknown, path: string, min: number, max: number, integer = false) {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < min || v > max || (integer && !Number.isInteger(v)))
    throw new InputError(path, `expected ${integer ? 'an integer' : 'a number'} from ${min} to ${max}, got ${what(v)}`)
}

/** The lines of a border (`BorderLine`), as xlsx names them; the xlsx reader takes these and names any other lost. */
export const LINES: readonly BorderLine['style'][] = ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']

/** A cell's alignments (`AlignmentStyle`), as xlsx names them. */
export const HORIZONTAL: readonly NonNullable<AlignmentStyle['horizontal']>[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
export const VERTICAL: readonly NonNullable<AlignmentStyle['vertical']>[] = ['top', 'center', 'bottom', 'justify', 'distributed']

/** Whether a field of a style is given: an undefined one never is, as no field of input is; in a patch a null one neither (`checkStyle`). */
export type Given = (o: Record<string, unknown>, k: string) => boolean

/** A font's fields: a cell's (`named`) has a name and a size too, which a conditional style (a dxf) has not. */
export function checkFont(v: unknown, p: string, given: Given, named: boolean) {
  const f = object(v, p, 'a font such as { bold: true }')
  onlyKeys(f, p, [...(named ? ['name', 'size'] : []), 'bold', 'italic', 'underline', 'strike', 'color'])
  if (given(f, 'name')) checkString(f.name, `${p}.name`)
  if (given(f, 'size')) checkNumber(f.size, `${p}.size`, 1, 409)
  for (const k of ['bold', 'italic', 'strike']) if (given(f, k)) checkBoolean(f[k], `${p}.${k}`)
  if (given(f, 'underline')) checkOneOf(f.underline, `${p}.underline`, ['single', 'double'])
  if (given(f, 'color')) checkColor(f.color, `${p}.color`)
}

/** A solid fill; `example` is the colour its message shows. */
export function checkFill(v: unknown, p: string, example: string) {
  const f = object(v, p, `a fill such as { color: "${example}" }`)
  onlyKeys(f, p, ['color'])
  checkColor(f.color, `${p}.color`)
}

/** A border's sides; `side` says which of those it names are given. */
export function checkBorder(v: unknown, p: string, side: Given) {
  const b = object(v, p, 'a border such as { bottom: { style: "thin" } }')
  onlyKeys(b, p, ['top', 'right', 'bottom', 'left'])
  for (const [k, l] of Object.entries(b)) {
    if (!side(b, k)) continue
    const lp = `${p}.${k}`
    const line = object(l, lp, 'a line such as { style: "thin", color: "#000000" }')
    onlyKeys(line, lp, ['style', 'color'])
    checkOneOf(line.style, `${lp}.style`, LINES)
    if (line.color !== undefined) checkColor(line.color, `${lp}.color`)
  }
}

/** A style, or with `patch` the fields of one to change (`StylePatch`), where null removes a field. */
export function checkStyle(v: unknown, p: string, patch = false): CellStyle {
  const s = object(v, p, 'a style such as { font: { bold: true } }')
  onlyKeys(s, p, ['font', 'fill', 'border', 'alignment'])
  // a field given: in a patch, null is a field removed
  const given: Given = (o, k) => o[k] !== undefined && !(patch && o[k] === null)
  if (given(s, 'font')) checkFont(s.font, `${p}.font`, given, true)
  if (given(s, 'fill')) checkFill(s.fill, `${p}.fill`, '#FFC000')
  if (given(s, 'border')) checkBorder(s.border, `${p}.border`, given)
  if (given(s, 'alignment')) {
    const a = object(s.alignment, `${p}.alignment`, 'an alignment such as { horizontal: "center" }')
    onlyKeys(a, `${p}.alignment`, ['horizontal', 'vertical', 'wrap', 'indent'])
    if (given(a, 'horizontal')) checkOneOf(a.horizontal, `${p}.alignment.horizontal`, HORIZONTAL)
    if (given(a, 'vertical')) checkOneOf(a.vertical, `${p}.alignment.vertical`, VERTICAL)
    if (given(a, 'wrap')) checkBoolean(a.wrap, `${p}.alignment.wrap`)
    if (given(a, 'indent')) checkNumber(a.indent, `${p}.alignment.indent`, 0, 250, true)
  }
  return s as CellStyle
}

export function checkStyles(v: unknown, path: string): CellStyle[] {
  return array(v, path, 'an array of styles').map((s, i) => checkStyle(s, `${path}[${i}]`))
}

const LINE = { c: { size: 'width', max: 255, what: 'columns', key: 'letters such as "C" or a span such as "B:D"' }, r: { size: 'height', max: 409, what: 'rows', key: 'a row number such as "5" or a span such as "5:9"' } } as const

/** A column or row entry, or a patch of one (fields may be null). */
function checkLine(v: unknown, p: string, axis: Axis, styles: number, patch: boolean) {
  const { size, max } = LINE[axis]
  const o = object(v, p, `an object such as { ${size}: 20 }`)
  onlyKeys(o, p, [size, ...(axis === 'c' ? ['bestFit'] : []), 'style', 'numberFormat', 'hidden', 'outlineLevel', 'collapsed'])
  const field = (k: string, check: (x: unknown, path: string) => void) => o[k] !== undefined && !(patch && o[k] === null) && check(o[k], `${p}.${k}`)
  field(size, (x, path) => checkNumber(x, path, 0, max))
  field('style', (x, path) => checkStyleIndex(x, path, styles))
  field('numberFormat', checkFormat)
  field('hidden', checkBoolean)
  field('outlineLevel', (x, path) => checkNumber(x, path, 1, 7, true))
  field('collapsed', checkBoolean)
  field('bestFit', checkBoolean)
}

/** The lines of a key (`C`, `B:D`; `5`, `5:9`). */
/** The lines of a key of `columns` or `rows` (`C`, `B:D`; `5`, `5:9`), else `InputError` naming `path`. */
export function checkSpan(key: unknown, path: string, axis: Axis): [number, number] {
  const sp = typeof key === 'string' ? parseSpan(key, axis) : null
  if (!sp) throw new InputError(path, `expected ${LINE[axis].key} within the sheet, got ${what(key)}`)
  return sp
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

/** Iterative calculation's bounds (`BookInput.iteration`): 1 to 32,767 rounds, as Excel takes them, a change from 0. */
export const ITERATION = { maxIterations: [1, 32_767], maxChange: [0, Number.MAX_VALUE] } as const

/** A hidden sheet's states (`SheetInput.hidden`) with the names an xlsx file gives them (`<sheet state>`). */
export const SHEET_STATES = new Map<true | 'very', string>([
  [true, 'hidden'],
  ['very', 'veryHidden'],
])

/**
 * The book: sheets with unique names by Excel's rules (compared case-insensitively), one of them visible at least, cells
 * by A1 address, columns, rows, merged areas, frozen panes, data validation and outline settings, names, styles. It is
 * returned as the engine takes it: its formulas without a leading `=`, its number formats by the name `numberFormat` (a
 * copy of what needs it, `taken.ts`).
 */
export function checkBook(v: unknown, path = 'book'): BookInput {
  const book = object(v, path, 'a book such as { sheets: [{ name: "Sheet1", cells: {} }] }')
  onlyKeys(book, path, ['sheets', 'names', 'styles', 'locale', 'dateSystem', 'iteration', 'externalBooks'])
  if (book.locale !== undefined) checkBookLocale(book.locale, `${path}.locale`)
  if (book.dateSystem !== undefined) checkOneOf(book.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (book.iteration !== undefined) {
    const it = object(book.iteration, `${path}.iteration`, 'iteration settings such as { maxIterations: 100, maxChange: 0.001 }')
    onlyKeys(it, `${path}.iteration`, ['maxIterations', 'maxChange'])
    if (it.maxIterations !== undefined) checkNumber(it.maxIterations, `${path}.iteration.maxIterations`, ...ITERATION.maxIterations, true)
    if (it.maxChange !== undefined) checkNumber(it.maxChange, `${path}.iteration.maxChange`, ...ITERATION.maxChange)
  }
  const styles = book.styles === undefined ? 0 : checkStyles(book.styles, `${path}.styles`).length
  const seen = new Map<string, { name: string; i: number }>()
  // the cells of each sheet that the engine takes otherwise: a formula with a leading `=`
  const taken: string[][] = []
  array(book.sheets, `${path}.sheets`, 'an array of sheets').forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(s, p, 'a sheet such as { name: "Sheet1", cells: {} }')
    onlyKeys(sheet, p, ['name', 'cells', 'columns', 'rows', 'merges', 'filtered', 'tables', 'frozen', 'conditionalFormats', 'validations', 'outline', 'hidden'])
    const name = checkString(sheet.name, `${p}.name`)
    const problem = sheetNameProblem(name)
    if (problem) throw new InputError(`${p}.name`, english(problem))
    const dup = seen.get(sheetKey(name))
    if (dup) throw new InputError(`${p}.name`, `duplicate sheet name ${JSON.stringify(name)}: ${path}.sheets[${dup.i}] is ${JSON.stringify(dup.name)} (sheet names are case-insensitive)`)
    seen.set(sheetKey(name), { name, i })
    const cells = object(sheet.cells, `${p}.cells`, 'an object of cells by address such as { A1: { value: 1 } }')
    // the cells of old-kind array formulas, checked against the sheet once all its cells are
    const arrays: string[] = []
    for (const addr of Object.keys(cells)) {
      const cell = cells[addr]
      if (SURE_ADDRESS.test(addr) && plainCell(cell, styles)) continue
      const cp = `${p}.cells[${JSON.stringify(addr)}]`
      checkAddress(addr, cp)
      const c = checkCell(cell, cp, styles)
      if (c.array !== undefined) arrays.push(addr)
      if (leadingEquals(c.formula)) (taken[i] ??= []).push(addr)
    }
    if (arrays.length) checkArrays(cells as Record<string, CellInput>, arrays, p)
    if (sheet.columns !== undefined) checkLines(sheet.columns, `${p}.columns`, 'c', styles)
    if (sheet.rows !== undefined) checkLines(sheet.rows, `${p}.rows`, 'r', styles)
    if (sheet.merges !== undefined) checkMerges(sheet.merges, `${p}.merges`)
    if (sheet.filtered !== undefined) checkBoolean(sheet.filtered, `${p}.filtered`)
    if (sheet.frozen !== undefined) checkFrozen(sheet.frozen, `${p}.frozen`)
    if (sheet.conditionalFormats !== undefined) checkConditionalFormats(sheet.conditionalFormats, `${p}.conditionalFormats`)
    if (sheet.validations !== undefined) checkValidations(sheet.validations, `${p}.validations`)
    if (sheet.outline !== undefined) checkOutline(sheet.outline, `${p}.outline`)
    if (sheet.hidden !== undefined && sheet.hidden !== false && !SHEET_STATES.has(sheet.hidden as true)) throw new InputError(`${p}.hidden`, `expected true, false or "very", got ${what(sheet.hidden)}`)
  })
  // Excel keeps one visible sheet at least
  const sheets = book.sheets as { hidden?: unknown }[]
  if (sheets.length && sheets.every((s) => s.hidden !== undefined && s.hidden !== false)) throw new InputError(`${path}.sheets`, 'a book keeps one visible sheet at least: every sheet is hidden')
  if (book.names !== undefined) checkNames(book.names, `${path}.names`)
  if (book.externalBooks !== undefined) checkExternalBooks(book.externalBooks, `${path}.externalBooks`)
  checkTables(book as unknown as BookInput, path)
  return bookTaken(book as unknown as BookInput, taken)
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

export function checkClock(v: unknown, path: string): Clock {
  const c = object(v, path, 'a clock such as { now: "2026-01-31T12:00:00Z", seed: 1 }')
  onlyKeys(c, path, ['now', 'seed'])
  const now = checkString(c.now, `${path}.now`)
  if (Number.isNaN(Date.parse(now))) throw new InputError(`${path}.now`, `invalid moment ${JSON.stringify(now)}: expected ISO 8601 such as "2026-01-31T12:00:00Z"`)
  if (typeof c.seed !== 'number' || !Number.isFinite(c.seed)) throw new InputError(`${path}.seed`, `expected a finite number, got ${what(c.seed)}`)
  return c as unknown as Clock
}

function checkExternals(v: unknown, path: string): Externals {
  const e = object(v, path, 'externals such as { functions: ["BALANCE"], values: {} }')
  onlyKeys(e, path, ['functions', 'values'])
  array(e.functions, `${path}.functions`, 'an array of function names').forEach((f, i) => checkString(f, `${path}.functions[${i}]`))
  if (e.values !== undefined) object(e.values, `${path}.values`, 'an object of values by call text')
  return e as unknown as Externals
}

const EXECUTIONS: readonly Execution[] = ['closures', 'eval']

function checkExecution(v: unknown, path: string) {
  if (!EXECUTIONS.includes(v as Execution)) throw new InputError(path, `expected "closures" or "eval", got ${what(v)}`)
}

const CLOCK_HINT = ' (pass the clock as { clock: { now, seed } })'

/** `limits`: whole numbers of at least 1. */
export function checkLimits(v: unknown, path: string): CalcLimits {
  const o = object(v, path, 'limits such as { formulas: 100000, cells: 10000000 }')
  onlyKeys(o, path, ['formulas', 'cells', 'spill'])
  for (const k of ['formulas', 'cells', 'spill']) if (o[k] !== undefined) checkNumber(o[k], `${path}.${k}`, 1, Number.MAX_SAFE_INTEGER, true)
  return o as CalcLimits
}

export function checkComputeOptions(v: unknown, path = 'opts'): ComputeOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } }')
  onlyOptions(o, path, ['clock', 'externals', 'execution', 'display', 'limits'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.clock !== undefined) checkClock(o.clock, `${path}.clock`)
  if (o.limits !== undefined) checkLimits(o.limits, `${path}.limits`)
  if (o.externals !== undefined) checkExternals(o.externals, `${path}.externals`)
  if (o.execution !== undefined) checkExecution(o.execution, `${path}.execution`)
  if (o.display !== undefined) checkBoolean(o.display, `${path}.display`)
  return o as ComputeOptions
}

export function checkSessionOptions(v: unknown, path = 'opts'): SessionOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } } or { live: true }')
  onlyOptions(o, path, ['clock', 'live', 'externals', 'execution', 'display', 'limits', 'undo'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.live !== undefined) checkBoolean(o.live, `${path}.live`)
  if (o.limits !== undefined) checkLimits(o.limits, `${path}.limits`)
  if (o.undo !== undefined) {
    const u = object(o.undo, `${path}.undo`, 'undo options such as { depth: 100, maxMb: 64 }')
    onlyKeys(u, `${path}.undo`, ['depth', 'maxMb'])
    if (u.depth !== undefined) checkNumber(u.depth, `${path}.undo.depth`, 0, 1_000_000, true)
    if (u.maxMb !== undefined) checkNumber(u.maxMb, `${path}.undo.maxMb`, 0, 1_000_000)
  }
  if (o.display !== undefined) checkBoolean(o.display, `${path}.display`)
  if (o.live === true && o.clock !== undefined) throw new InputError(path, 'a live session takes its clock from the system: pass either live or clock')
  if (o.clock !== undefined) checkClock(o.clock, `${path}.clock`)
  if (o.externals !== undefined) checkExternals(o.externals, `${path}.externals`)
  if (o.execution !== undefined) checkExecution(o.execution, `${path}.execution`)
  return o as SessionOptions
}

/** Edits: each names a sheet of the book (`sheetIndex` resolves it) and a valid address; the book has `styles` styles. */
/** Edits of cells as the engine takes them (`taken.ts`). */
export function checkEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): CellEdit[] {
  const edits = array(v, path, 'an array of edits such as [{ sheet: "Sheet1", cell: "A1", input: { value: 1 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'an edit such as { sheet: "Sheet1", cell: "A1", input: { value: 1 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'input'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    if (o.input === undefined) throw new InputError(`${p}.input`, 'expected a cell object, or null to clear the cell')
    if (o.input !== null) checkCell(o.input, `${p}.input`, styles)
    return o as unknown as CellEdit
  })
  return editsTaken(edits)
}

/**
 * Layout edits of a book with `styles` styles: a sheet of the book, and columns, rows, merged areas, the filter mode, frozen
 * panes, rules of data validation or outline settings; as the engine takes them (`taken.ts`).
 */
export function checkLayoutEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): LayoutEdit[] {
  const edits = array(v, path, 'an array of layout edits such as [{ sheet: "Sheet1", columns: "C", set: { width: 20 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'a layout edit such as { sheet: "Sheet1", columns: "C", set: { width: 20 } }')
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    if (o.merges !== undefined) {
      onlyKeys(o, p, ['sheet', 'merges'])
      checkMerges(o.merges, `${p}.merges`)
      return o as unknown as LayoutEdit
    }
    if (o.filtered !== undefined) {
      onlyKeys(o, p, ['sheet', 'filtered'])
      checkBoolean(o.filtered, `${p}.filtered`)
      return o as unknown as LayoutEdit
    }
    if ('frozen' in o) {
      onlyKeys(o, p, ['sheet', 'frozen'])
      if (o.frozen !== null) checkFrozen(o.frozen, `${p}.frozen`)
      return o as unknown as LayoutEdit
    }
    if ('conditionalFormats' in o) {
      onlyKeys(o, p, ['sheet', 'conditionalFormats'])
      checkConditionalFormats(o.conditionalFormats, `${p}.conditionalFormats`)
      return o as unknown as LayoutEdit
    }
    if ('validations' in o || 'outline' in o) {
      const k = 'validations' in o ? 'validations' : 'outline'
      onlyKeys(o, p, ['sheet', k])
      if (k === 'validations') checkValidations(o.validations, `${p}.validations`)
      else if (o.outline !== null) checkOutline(o.outline, `${p}.outline`)
      return o as unknown as LayoutEdit
    }
    if (o.fit !== undefined) {
      onlyKeys(o, p, ['sheet', 'columns', 'fit'])
      if (o.fit !== true) throw new InputError(`${p}.fit`, `expected true, got ${what(o.fit)}`)
      checkSpan(o.columns, `${p}.columns`, 'c')
      return o as unknown as LayoutEdit
    }
    const axis: Axis = o.rows !== undefined ? 'r' : 'c'
    const lines = axis === 'r' ? 'rows' : 'columns'
    onlyKeys(o, p, ['sheet', lines, 'set'], o.columns === undefined && o.rows === undefined ? ' (one of columns, rows, merges, filtered, frozen, conditionalFormats, validations or outline)' : '')
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
  if (o.locale !== undefined) checkLocaleArg(o.locale, `${path}.locale`)
  return o as NameOptions
}

/** `validateNames`'s context: the book's sheets, which a name's scope needs, its tables, external functions and locale. */
export function checkNamesContext(v: unknown, path = 'ctx'): NamesContext {
  const o = object(v, path, "a context such as { sheets: ['Data'], locale: 'ru-RU' }")
  onlyKeys(o, path, ['sheets', 'tables', 'externalFunctions', 'locale', 'externalBooks'])
  if (o.sheets === undefined) throw new InputError(`${path}.sheets`, "expected the book's sheet names, an array of strings")
  const { locale, ...rest } = o
  checkValidateContext(rest, path)
  if (locale !== undefined) checkLocaleArg(locale, `${path}.locale`)
  return o as unknown as NamesContext
}

/** `displayValue`'s value: a number, text, a boolean, null or an error `{ error: code }`. */
export function checkCellValue(v: unknown, path: string): CellValue {
  if (v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) return v
  if (typeof v === 'object' && !Array.isArray(v) && Object.keys(v).length === 1 && (v as { error?: unknown }).error as string in ERROR_TEXT) return v as CellValue
  throw new InputError(path, `expected a number, text, a boolean, null or { error: code }, got ${what(v)}`)
}

/** `displayValue`'s options. */
export function checkDisplayOptions(v: unknown, path = 'opts'): DisplayOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { numberFormat: '#,##0.00', locale: 'ru-RU', width: 12 }")
  onlyKeys(o, path, ['numberFormat', 'locale', 'width', 'dateSystem'])
  if (o.numberFormat !== undefined) checkFormat(o.numberFormat, `${path}.numberFormat`)
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (o.locale !== undefined) checkLocaleArg(o.locale, `${path}.locale`)
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

/** `entryText`'s and `needsApostrophe`'s options: the cell's number format, the book's locale and date system, the moment for a date without a year. */
export function checkEntryTextOptions(v: unknown, path = 'opts'): EntryTextOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { numberFormat: '0%', locale: 'ru-RU' }")
  onlyKeys(o, path, ['numberFormat', 'locale', 'dateSystem', 'now'])
  if (o.numberFormat !== undefined) checkFormat(o.numberFormat, `${path}.numberFormat`)
  checkEntryOptions({ ...(o.locale !== undefined ? { locale: o.locale } : {}), ...(o.dateSystem !== undefined ? { dateSystem: o.dateSystem } : {}), ...(o.now !== undefined ? { now: o.now } : {}) }, path)
  return o as EntryTextOptions
}

/**
 * `parseEntry`'s options; of the cell only its number format, style, meta, note and link are read (it may be a
 * `CellView`), and the old name of the number format is refused.
 */
export function checkEntryOptions(v: unknown, path = 'opts'): EntryOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { locale: 'ru-RU', cell: { numberFormat: '0.00' } }")
  onlyKeys(o, path, ['locale', 'dateSystem', 'cell', 'now', 'externalBooks'])
  if (o.locale !== undefined) checkLocaleArg(o.locale, `${path}.locale`)
  if (o.externalBooks !== undefined) checkExternalBooks(o.externalBooks, `${path}.externalBooks`, false)
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (o.now !== undefined && Number.isNaN(Date.parse(checkString(o.now, `${path}.now`)))) throw new InputError(`${path}.now`, `invalid moment ${JSON.stringify(o.now)}: expected ISO 8601 such as "2026-01-31T12:00:00Z"`)
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
