import { HyperFormula, AlwaysSparse, DenseSparseChooseBasedOnThreshold } from 'hyperformula'
import { compute, createSession } from '../../dist/index.js'
import { dateToSerial, parseAddress } from '../../dist/model.js'

export { HyperFormula }
export const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }
export const HF_OPTIONS = {
  licenseKey: 'gpl-v3', language: 'enGB', localeLang: 'en-US', accentSensitive: true,
  useArrayArithmetic: true, ignoreWhiteSpace: 'any', evaluateNullToZero: true,
  leapYear1900: true, nullDate: { year: 1899, month: 12, day: 31 },
  precisionRounding: 15, maxRows: 1_048_576, maxColumns: 16_384,
}

export const isHf = (engine) => engine.startsWith('hf-')
const inputValue = (cell) => {
  const value = cell?.value ?? null
  return value !== null && typeof value === 'object' ? dateToSerial(value.date) : value
}

export function hfScalar(cell) {
  if (cell?.formula !== undefined) return cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`
  const value = inputValue(cell)
  return typeof value === 'string' ? `'${value}` : value
}

export function hfInput(book) {
  const sheets = {}
  for (const sheet of book.sheets) {
    const rows = []
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      ;(rows[row] ??= [])[col] = hfScalar(cell)
    }
    sheets[sheet.name] = Array.from(rows, (row) => row ?? [])
  }
  const names = [
    { name: 'TRUE', expression: '=TRUE()' }, { name: 'FALSE', expression: '=FALSE()' },
    ...(book.names ?? []).map((n) => ({ name: n.name, expression: n.ref.startsWith('=') ? n.ref : `=${n.ref}`, ...(n.scope === undefined ? {} : { scope: book.sheets.findIndex((s) => s.name === n.scope) }) })),
  ]
  return { sheets, names }
}

export function open(engine, book, prepared) {
  return isHf(engine)
    ? HyperFormula.buildFromSheets(prepared.sheets, { ...HF_OPTIONS, useColumnIndex: engine === 'hf-index', ...(engine === 'hf-sparse' ? { chooseAddressMappingPolicy: new AlwaysSparse() } : engine === 'hf-auto' ? { chooseAddressMappingPolicy: new DenseSparseChooseBasedOnThreshold(0.5) } : {}) }, prepared.names)
    : createSession(book, { clock: CLOCK, execution: engine === 'xl-eval' ? 'eval' : 'closures', display: false })
}

export function prepareEdit(engine, session, edits) {
  if (!isHf(engine)) return () => session.setCells(edits)
  const calls = edits.map((e) => ({ address: { sheet: session.getSheetId(e.sheet), ...parseAddress(e.cell) }, values: [[hfScalar(e.input)]] }))
  const first = calls[0]
  if (first && calls.every((x, i) => x.address.sheet === first.address.sheet && x.address.col === first.address.col && x.address.row === first.address.row + i)) {
    const values = calls.map((x) => x.values[0])
    return () => session.setCellContents(first.address, values)
  }
  return () => session.batch(() => { for (const c of calls) session.setCellContents(c.address, c.values) })
}

export const close = (engine, session) => isHf(engine) ? session.destroy() : session.close()

export function reference(book, edits = []) {
  const grouped = new Map()
  for (const e of edits) {
    let cells = grouped.get(e.sheet)
    if (!cells) grouped.set(e.sheet, (cells = {}))
    cells[e.cell] = e.input
  }
  const changed = { ...book, sheets: book.sheets.map((s) => grouped.has(s.name) ? { ...s, cells: { ...s.cells, ...grouped.get(s.name) } } : s) }
  const result = compute(changed, { clock: CLOCK, display: false })
  // Keep computed outputs only, instead of duplicating half a million unchanged literal cells for each edit.
  const formulas = new Map()
  const sourceSheets = new Map(book.sheets.map((s) => [s.name, s.cells]))
  for (const sheet of result.sheets) for (const [addr, cell] of Object.entries(sheet.cells))
    if (sourceSheets.get(sheet.name)[addr]?.formula !== undefined) formulas.set(`${sheet.name}!${addr}`, cell.value)
  return { formulas, grouped, hash: result.hash, inputHash: result.inputHash }
}

const ERRORS = { DIV_BY_ZERO: 'DIV0', NA: 'NA', NAME: 'NAME', NUM: 'NUM', REF: 'REF', VALUE: 'VALUE', CYCLE: 'CYCLE', SPILL: 'SPILL', ERROR: 'ERROR' }
const normalize = (value) => value && typeof value === 'object' && 'type' in value ? { error: ERRORS[value.type] ?? value.type } : value
const equal = (a, b) => typeof a === 'number' && typeof b === 'number'
  ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b))
  : JSON.stringify(a) === JSON.stringify(b)

export function verify(engine, session, book, expected) {
  const actual = isHf(engine) ? null : new Map(session.getResult().sheets.map((s) => [s.name, s.cells]))
  const mismatches = []
  let checked = 0, maxRelativeError = 0, differences = 0
  for (const sheet of book.sheets) {
    const id = isHf(engine) ? session.getSheetId(sheet.name) : undefined
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      const a = cell.formula !== undefined ? expected.formulas.get(`${sheet.name}!${addr}`) : inputValue(expected.grouped.get(sheet.name)?.[addr] ?? cell)
      const b = actual ? actual.get(sheet.name)[addr]?.value ?? null : normalize(session.getCellValue({ sheet: id, ...parseAddress(addr) }))
      checked++
      if (typeof a === 'number' && typeof b === 'number') maxRelativeError = Math.max(maxRelativeError, Math.abs(a - b) / Math.max(1, Math.abs(a), Math.abs(b)))
      if (!equal(a, b)) {
        differences++
        if (mismatches.length < 10) mismatches.push({ sheet: sheet.name, cell: addr, expected: a, actual: b })
      }
    }
  }
  return { checked, differences, maxRelativeError, mismatches }
}

export function measure(f) {
  const t = performance.now(), cpu = process.cpuUsage()
  const value = f()
  const wallMs = performance.now() - t, usage = process.cpuUsage(cpu)
  return { value, wallMs, cpuMs: (usage.user + usage.system) / 1000 }
}

export const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)]
export const p95 = (values) => [...values].sort((a, b) => a - b)[Math.ceil(values.length * 0.95) - 1]
