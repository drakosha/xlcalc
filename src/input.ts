/**
 * Checks of caller input at the API boundary. Wrong input (a value of the wrong type, a malformed cell
 * address, duplicate sheets, an unknown sheet in an edit) throws `InputError` naming the place and the
 * problem. Formula problems are not input errors: an invalid formula gives an error value in its cell.
 */
import type { BookInput, CalcLimits, CellEdit, CellInput, CellStyle, CellValue, Clock, ComputeOptions, DisplayOptions, Execution, Externals, LayoutEdit, NameInput, SessionOptions, ValidateContext } from './contract.js'
import { overlapping, parseMerge } from './layout/merges.js'
import { parseSpan } from './layout/spans.js'
import { LOCALES } from './locale/registry.js'
import { dateToSerial, MAX_COLS, MAX_ROWS } from './model.js'
import { isBuiltin } from './numfmt/builtin.js'
import type { Axis } from './structure/area.js'
import { checkArrays } from './array-areas.js'
import { ERROR_TEXT, sheetKey } from './syntax.js'
import { checkTables } from './tables/check.js'

/** Input of the wrong shape. `path` points at the offending part, for example `book.sheets[1].cells["B0"]`. */
export class InputError extends TypeError {
  override readonly name = 'InputError'
  constructor(
    readonly path: string,
    problem: string,
  ) {
    super(`${path}: ${problem}`)
  }
}

function what(v: unknown): string {
  if (v === undefined || v === null) return String(v)
  if (Array.isArray(v)) return 'an array'
  if (typeof v === 'object') return 'an object'
  if (typeof v === 'function') return 'a function'
  if (typeof v === 'string') return `the string ${JSON.stringify(v)}`
  return `${typeof v} ${String(v)}`
}

const isObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v)

function object(v: unknown, path: string, expected: string): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

function array(v: unknown, path: string, expected: string): unknown[] {
  if (!Array.isArray(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

export function checkString(v: unknown, path: string): string {
  if (typeof v !== 'string') throw new InputError(path, `expected a string, got ${what(v)}`)
  return v
}

function onlyKeys(o: Record<string, unknown>, path: string, allowed: readonly string[], hint = '') {
  for (const k of Object.keys(o)) if (!allowed.includes(k)) throw new InputError(path, `unknown field ${JSON.stringify(k)} (expected ${allowed.join(', ')})${hint}`)
}

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
function checkFormat(v: unknown, path: string) {
  if (typeof v !== 'string' && !(typeof v === 'number' && isBuiltin(v))) throw new InputError(path, `expected a format code or the number of a built-in format, 0 to 49, got ${what(v)}`)
}

/** A style number of a book with `styles` styles. */
function checkStyleIndex(v: unknown, path: string, styles: number) {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v >= styles)
    throw new InputError(path, styles ? `expected the index of a style of the book, 0 to ${styles - 1}, got ${what(v)}` : `the book has no styles, got ${what(v)}`)
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
  if (c.format !== undefined && typeof c.format !== 'string') checkFormat(c.format, `${path}.format`)
  if (c.style !== undefined) checkStyleIndex(c.style, `${path}.style`, styles)
  if (c.array !== undefined) checkString(c.array, `${path}.array`)
  return c as CellInput
}

export function checkNames(v: unknown, path: string): NameInput[] {
  return array(v, path, 'an array of names').map((n, i) => {
    const p = `${path}[${i}]`
    const o = object(n, p, "a name such as { name: 'rate', ref: \"'Data'!$B$1\" }")
    onlyKeys(o, p, ['name', 'ref', 'scope'])
    checkString(o.name, `${p}.name`)
    checkString(o.ref, `${p}.ref`)
    if (o.scope !== undefined) checkString(o.scope, `${p}.scope`)
    return o as unknown as NameInput
  })
}

/** Within the grid for sure (up to ZZ999999); anything else takes the full check. */
const SURE_ADDRESS = /^[A-Z]{1,2}[1-9][0-9]{0,5}$/

const CELL_FIELDS = ['value', 'formula', 'format', 'style', 'meta', 'array']

/** The common cell, checked without building paths (the cell loop is hot); anything else takes the full check. */
function plainCell(c: unknown, styles: number): boolean {
  if (!isObject(c)) return false
  for (const k in c) {
    const v = c[k]
    if (k === 'meta') continue
    if (k === 'value' && (v === null || v === undefined || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean')) continue
    if ((k === 'formula' || k === 'format') && (v === undefined || typeof v === 'string')) continue
    if (k === 'style' && Number.isInteger(v) && (v as number) >= 0 && (v as number) < styles) continue
    return false
  }
  return true
}

/** Why Excel would not take a sheet name, or null: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` (reserved, in any case). */
export function sheetNameProblem(name: string): string | null {
  if (name === '') return 'a sheet name must not be empty'
  if (name.length > 31) return `a sheet name has at most 31 characters, ${JSON.stringify(name)} has ${name.length}`
  const bad = /[:\\/?*[\]]/.exec(name)
  if (bad) return `a sheet name must not contain ${JSON.stringify(bad[0])} (Excel forbids : \\ / ? * [ ])`
  if (name.startsWith("'") || name.endsWith("'")) return 'a sheet name must not begin or end with an apostrophe'
  // Excel refuses to rename a sheet so (the struct set `history`)
  if (name.toLowerCase() === 'history') return 'Excel reserves the sheet name "History" (in any case)'
  return null
}

const COLOR = /^#[0-9A-Fa-f]{6}$/

function checkColor(v: unknown, path: string) {
  if (typeof v !== 'string' || !COLOR.test(v)) throw new InputError(path, `expected a colour such as "#FF0000", got ${what(v)}`)
}

function checkOneOf(v: unknown, path: string, values: readonly (string | number)[]) {
  if (!values.includes(v as string)) throw new InputError(path, `expected one of ${values.map((x) => JSON.stringify(x)).join(', ')}, got ${what(v)}`)
}

const DATE_SYSTEMS = [1900, 1904]

function checkBoolean(v: unknown, path: string) {
  if (typeof v !== 'boolean') throw new InputError(path, `expected a boolean, got ${what(v)}`)
}

function checkNumber(v: unknown, path: string, min: number, max: number, integer = false) {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < min || v > max || (integer && !Number.isInteger(v)))
    throw new InputError(path, `expected ${integer ? 'an integer' : 'a number'} from ${min} to ${max}, got ${what(v)}`)
}

const LINES = ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']

function checkStyle(v: unknown, p: string): CellStyle {
  const s = object(v, p, 'a style such as { font: { bold: true } }')
  onlyKeys(s, p, ['font', 'fill', 'border', 'alignment'])
  if (s.font !== undefined) {
    const f = object(s.font, `${p}.font`, 'a font such as { bold: true }')
    onlyKeys(f, `${p}.font`, ['name', 'size', 'bold', 'italic', 'underline', 'strike', 'color'])
    if (f.name !== undefined) checkString(f.name, `${p}.font.name`)
    if (f.size !== undefined) checkNumber(f.size, `${p}.font.size`, 1, 409)
    for (const k of ['bold', 'italic', 'strike']) if (f[k] !== undefined) checkBoolean(f[k], `${p}.font.${k}`)
    if (f.underline !== undefined) checkOneOf(f.underline, `${p}.font.underline`, ['single', 'double'])
    if (f.color !== undefined) checkColor(f.color, `${p}.font.color`)
  }
  if (s.fill !== undefined) {
    const f = object(s.fill, `${p}.fill`, 'a fill such as { color: "#FFC000" }')
    onlyKeys(f, `${p}.fill`, ['color'])
    checkColor(f.color, `${p}.fill.color`)
  }
  if (s.border !== undefined) {
    const b = object(s.border, `${p}.border`, 'a border such as { bottom: { style: "thin" } }')
    onlyKeys(b, `${p}.border`, ['top', 'right', 'bottom', 'left'])
    for (const [side, l] of Object.entries(b)) {
      const lp = `${p}.border.${side}`
      const line = object(l, lp, 'a line such as { style: "thin", color: "#000000" }')
      onlyKeys(line, lp, ['style', 'color'])
      checkOneOf(line.style, `${lp}.style`, LINES)
      if (line.color !== undefined) checkColor(line.color, `${lp}.color`)
    }
  }
  if (s.alignment !== undefined) {
    const a = object(s.alignment, `${p}.alignment`, 'an alignment such as { horizontal: "center" }')
    onlyKeys(a, `${p}.alignment`, ['horizontal', 'vertical', 'wrap', 'indent'])
    if (a.horizontal !== undefined) checkOneOf(a.horizontal, `${p}.alignment.horizontal`, ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'])
    if (a.vertical !== undefined) checkOneOf(a.vertical, `${p}.alignment.vertical`, ['top', 'center', 'bottom', 'justify', 'distributed'])
    if (a.wrap !== undefined) checkBoolean(a.wrap, `${p}.alignment.wrap`)
    if (a.indent !== undefined) checkNumber(a.indent, `${p}.alignment.indent`, 0, 250, true)
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
  onlyKeys(o, p, [size, 'style', 'format', 'hidden'])
  const field = (k: string, check: (x: unknown, path: string) => void) => o[k] !== undefined && !(patch && o[k] === null) && check(o[k], `${p}.${k}`)
  field(size, (x, path) => checkNumber(x, path, 0, max))
  field('style', (x, path) => checkStyleIndex(x, path, styles))
  field('format', checkFormat)
  field('hidden', checkBoolean)
}

/** The lines of a key (`C`, `B:D`; `5`, `5:9`). */
function checkSpan(key: unknown, path: string, axis: Axis): [number, number] {
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

/**
 * The book: sheets with unique names by Excel's rules (compared case-insensitively), cells by A1 address,
 * columns, rows and merged areas, names, styles.
 */
export function checkBook(v: unknown, path = 'book'): BookInput {
  const book = object(v, path, 'a book such as { sheets: [{ name: "Sheet1", cells: {} }] }')
  onlyKeys(book, path, ['sheets', 'names', 'styles', 'locale', 'dateSystem', 'iteration'])
  if (book.locale !== undefined) checkLocale(book.locale, `${path}.locale`)
  if (book.dateSystem !== undefined) checkOneOf(book.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (book.iteration !== undefined) {
    const it = object(book.iteration, `${path}.iteration`, 'iteration settings such as { maxIterations: 100, maxChange: 0.001 }')
    onlyKeys(it, `${path}.iteration`, ['maxIterations', 'maxChange'])
    if (it.maxIterations !== undefined) checkNumber(it.maxIterations, `${path}.iteration.maxIterations`, 1, 32_767, true)
    if (it.maxChange !== undefined) checkNumber(it.maxChange, `${path}.iteration.maxChange`, 0, Number.MAX_VALUE)
  }
  const styles = book.styles === undefined ? 0 : checkStyles(book.styles, `${path}.styles`).length
  const seen = new Map<string, { name: string; i: number }>()
  array(book.sheets, `${path}.sheets`, 'an array of sheets').forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(s, p, 'a sheet such as { name: "Sheet1", cells: {} }')
    onlyKeys(sheet, p, ['name', 'cells', 'columns', 'rows', 'merges', 'filtered', 'tables'])
    const name = checkString(sheet.name, `${p}.name`)
    const problem = sheetNameProblem(name)
    if (problem) throw new InputError(`${p}.name`, problem)
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
      if (checkCell(cell, cp, styles).array !== undefined) arrays.push(addr)
    }
    if (arrays.length) checkArrays(cells as Record<string, CellInput>, arrays, p)
    if (sheet.columns !== undefined) checkLines(sheet.columns, `${p}.columns`, 'c', styles)
    if (sheet.rows !== undefined) checkLines(sheet.rows, `${p}.rows`, 'r', styles)
    if (sheet.merges !== undefined) checkMerges(sheet.merges, `${p}.merges`)
    if (sheet.filtered !== undefined) checkBoolean(sheet.filtered, `${p}.filtered`)
  })
  if (book.names !== undefined) checkNames(book.names, `${path}.names`)
  checkTables(book as unknown as BookInput, path)
  return book as unknown as BookInput
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
  onlyKeys(o, path, ['formulas', 'cells'])
  for (const k of ['formulas', 'cells']) if (o[k] !== undefined) checkNumber(o[k], `${path}.${k}`, 1, Number.MAX_SAFE_INTEGER, true)
  return o as CalcLimits
}

/** A tag of the package's locales. */
function checkLocale(v: unknown, path: string) {
  if (!LOCALES.has(checkString(v, path))) throw new InputError(path, `unknown locale ${JSON.stringify(v)} (expected ${[...LOCALES.keys()].join(', ')})`)
}

export function checkComputeOptions(v: unknown, path = 'opts'): ComputeOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } }')
  onlyKeys(o, path, ['clock', 'externals', 'execution', 'display', 'limits'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
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
  onlyKeys(o, path, ['clock', 'live', 'externals', 'execution', 'display', 'limits'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.live !== undefined) checkBoolean(o.live, `${path}.live`)
  if (o.limits !== undefined) checkLimits(o.limits, `${path}.limits`)
  if (o.display !== undefined) checkBoolean(o.display, `${path}.display`)
  if (o.live === true && o.clock !== undefined) throw new InputError(path, 'a live session takes its clock from the system: pass either live or clock')
  if (o.clock !== undefined) checkClock(o.clock, `${path}.clock`)
  if (o.externals !== undefined) checkExternals(o.externals, `${path}.externals`)
  if (o.execution !== undefined) checkExecution(o.execution, `${path}.execution`)
  return o as SessionOptions
}

/** Edits: each names a sheet of the book (`sheetIndex` resolves it) and a valid address; the book has `styles` styles. */
export function checkEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): CellEdit[] {
  return array(v, path, 'an array of edits such as [{ sheet: "Sheet1", cell: "A1", input: { value: 1 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'an edit such as { sheet: "Sheet1", cell: "A1", input: { value: 1 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'input'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    if (o.input === undefined) throw new InputError(`${p}.input`, 'expected a cell object, or null to clear the cell')
    if (o.input !== null) checkCell(o.input, `${p}.input`, styles)
    return o as unknown as CellEdit
  })
}

/** Layout edits of a book with `styles` styles: a sheet of the book, and columns, rows or merged areas. */
export function checkLayoutEdits(v: unknown, sheetIndex: (name: string) => number | undefined, styles: number, path = 'edits'): LayoutEdit[] {
  return array(v, path, 'an array of layout edits such as [{ sheet: "Sheet1", columns: "C", set: { width: 20 } }]').map((e, i) => {
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
    const axis: Axis = o.rows !== undefined ? 'r' : 'c'
    const lines = axis === 'r' ? 'rows' : 'columns'
    onlyKeys(o, p, ['sheet', lines, 'set'], o.columns === undefined && o.rows === undefined ? ' (one of columns, rows, merges or filtered)' : '')
    checkSpan(o[lines], `${p}.${lines}`, axis)
    checkLine(o.set, `${p}.set`, axis, styles, true)
    return o as unknown as LayoutEdit
  })
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
  onlyKeys(o, path, ['sheets', 'names', 'externalFunctions', 'tables'])
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


/** `displayValue`'s value: a number, text, a boolean, null or an error `{ error: code }`. */
export function checkCellValue(v: unknown, path: string): CellValue {
  if (v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) return v
  if (typeof v === 'object' && !Array.isArray(v) && Object.keys(v).length === 1 && (v as { error?: unknown }).error as string in ERROR_TEXT) return v as CellValue
  throw new InputError(path, `expected a number, text, a boolean, null or { error: code }, got ${what(v)}`)
}

export function checkDisplayOptions(v: unknown, path = 'opts'): DisplayOptions {
  if (v === undefined) return {}
  const o = object(v, path, "options such as { format: '#,##0.00', locale: 'ru-RU', width: 12 }")
  onlyKeys(o, path, ['format', 'locale', 'width', 'dateSystem'])
  if (o.format !== undefined) checkFormat(o.format, `${path}.format`)
  if (o.dateSystem !== undefined) checkOneOf(o.dateSystem, `${path}.dateSystem`, DATE_SYSTEMS)
  if (o.locale !== undefined) checkLocale(o.locale, `${path}.locale`)
  if (o.width !== undefined) checkNumber(o.width, `${path}.width`, 0, 255 * MAX_COLS)
  return o as DisplayOptions
}
