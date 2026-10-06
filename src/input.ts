/**
 * Checks of caller input at the API boundary. Wrong input (a value of the wrong type, a malformed cell
 * address, duplicate sheets, an unknown sheet in an edit) throws `InputError` naming the place and the
 * problem. Formula problems are not input errors: an invalid formula gives an error value in its cell.
 */
import type { BookInput, CellEdit, CellInput, Clock, ComputeOptions, Execution, Externals, NameInput, SessionOptions, ValidateContext } from './contract.js'
import { dateToSerial, MAX_COLS, MAX_ROWS } from './model.js'
import { sheetKey } from './syntax.js'

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

export function checkCell(v: unknown, path: string): CellInput {
  const c = object(v, path, 'a cell object such as { value: 1 } or { formula: "A1*2" }')
  for (const k in c) if (k !== 'value' && k !== 'formula' && k !== 'format' && k !== 'meta') onlyKeys(c, path, ['value', 'formula', 'format', 'meta'])
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
  if (c.format !== undefined && typeof c.format !== 'string') checkString(c.format, `${path}.format`)
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

/** The common cell, checked without building paths (the cell loop is hot); anything else takes the full check. */
function plainCell(c: unknown): boolean {
  if (!isObject(c)) return false
  for (const k in c) {
    const v = c[k]
    if (k === 'meta') continue
    if (k === 'value' && (v === null || v === undefined || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean')) continue
    if ((k === 'formula' || k === 'format') && (v === undefined || typeof v === 'string')) continue
    return false
  }
  return true
}

/** The book: sheets with unique names (case-insensitive, as in Excel), cells by A1 address, names. */
export function checkBook(v: unknown, path = 'book'): BookInput {
  const book = object(v, path, 'a book such as { sheets: [{ name: "Sheet1", cells: {} }] }')
  const seen = new Map<string, { name: string; i: number }>()
  array(book.sheets, `${path}.sheets`, 'an array of sheets').forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    const sheet = object(s, p, 'a sheet such as { name: "Sheet1", cells: {} }')
    const name = checkString(sheet.name, `${p}.name`)
    if (name === '') throw new InputError(`${p}.name`, 'a sheet name must not be empty')
    const dup = seen.get(sheetKey(name))
    if (dup) throw new InputError(`${p}.name`, `duplicate sheet name ${JSON.stringify(name)}: ${path}.sheets[${dup.i}] is ${JSON.stringify(dup.name)} (sheet names are case-insensitive)`)
    seen.set(sheetKey(name), { name, i })
    const cells = object(sheet.cells, `${p}.cells`, 'an object of cells by address such as { A1: { value: 1 } }')
    for (const addr of Object.keys(cells)) {
      const cell = cells[addr]
      if (SURE_ADDRESS.test(addr) && plainCell(cell)) continue
      const cp = `${p}.cells[${JSON.stringify(addr)}]`
      checkAddress(addr, cp)
      checkCell(cell, cp)
    }
  })
  if (book.names !== undefined) checkNames(book.names, `${path}.names`)
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

export function checkComputeOptions(v: unknown, path = 'opts'): ComputeOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } }')
  onlyKeys(o, path, ['clock', 'externals', 'execution'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.clock !== undefined) checkClock(o.clock, `${path}.clock`)
  if (o.externals !== undefined) checkExternals(o.externals, `${path}.externals`)
  if (o.execution !== undefined) checkExecution(o.execution, `${path}.execution`)
  return o as ComputeOptions
}

export function checkSessionOptions(v: unknown, path = 'opts'): SessionOptions {
  if (v === undefined) return {}
  const o = object(v, path, 'options such as { clock: { now, seed } } or { live: true }')
  onlyKeys(o, path, ['clock', 'live', 'externals', 'execution'], 'now' in o || 'seed' in o ? CLOCK_HINT : '')
  if (o.live !== undefined && typeof o.live !== 'boolean') throw new InputError(`${path}.live`, `expected a boolean, got ${what(o.live)}`)
  if (o.live === true && o.clock !== undefined) throw new InputError(path, 'a live session takes its clock from the system: pass either live or clock')
  if (o.clock !== undefined) checkClock(o.clock, `${path}.clock`)
  if (o.externals !== undefined) checkExternals(o.externals, `${path}.externals`)
  if (o.execution !== undefined) checkExecution(o.execution, `${path}.execution`)
  return o as SessionOptions
}

/** Edits: each names a sheet of the book (`sheetIndex` resolves it) and a valid address. */
export function checkEdits(v: unknown, sheetIndex: (name: string) => number | undefined, path = 'edits'): CellEdit[] {
  return array(v, path, 'an array of edits such as [{ sheet: "Sheet1", cell: "A1", input: { value: 1 } }]').map((e, i) => {
    const p = `${path}[${i}]`
    const o = object(e, p, 'an edit such as { sheet: "Sheet1", cell: "A1", input: { value: 1 } }')
    onlyKeys(o, p, ['sheet', 'cell', 'input'])
    checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    checkAddress(o.cell, `${p}.cell`)
    if (o.input === undefined) throw new InputError(`${p}.input`, 'expected a cell object, or null to clear the cell')
    if (o.input !== null) checkCell(o.input, `${p}.input`)
    return o as unknown as CellEdit
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
  onlyKeys(o, path, ['sheets', 'names', 'externalFunctions'])
  for (const k of ['sheets', 'names', 'externalFunctions'] as const)
    if (o[k] !== undefined) array(o[k], `${path}.${k}`, 'an array of strings').forEach((s, i) => checkString(s, `${path}.${k}[${i}]`))
  return o as ValidateContext
}

