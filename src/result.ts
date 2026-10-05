/** Выдача расчёта по контракту: значения по листам, причины ошибок, волатильные ячейки, хэши. */
import { inputHash, outNumber, valuesHash } from './canon.js'
import { CONFIG_VERSION } from './config.js'
import type { CellResult, CellValue, Clock, ComputeResult } from './contract.js'
import { issues, sortedCells, toCellValue, valueAt, type EngineState } from './engine.js'
import { formatAddress, msToSerial, parseAddress } from './model.js'


/** Серийный момент часов (UTC); неверный момент это ошибка вызова. */
export function clockSerial(clock: Clock): number {
  const ms = Date.parse(clock.now)
  if (Number.isNaN(ms)) throw new Error(`bad clock.now: ${clock.now}`)
  return msToSerial(ms)
}

export function outValue(v: CellValue): CellValue {
  return typeof v === 'number' ? outNumber(v) : v
}

export function buildResult(state: EngineState, clock: Clock, t0: number): ComputeResult {
  const values = new Map<string, Map<string, CellValue>>()
  const sheets = state.input.sheets.map((sheet, s) => {
    const cells: Record<string, CellResult> = {}
    const vals = new Map<string, CellValue>()
    for (const addr of sortedCells(state, s)) {
      const { row, col } = parseAddress(addr)
      const inCell = sheet.cells[addr]
      const value = outValue(toCellValue(valueAt(state, s, row, col)))
      const id = state.at.get(`${s}:${row}:${col}`)
      const out: CellResult = { value }
      const text = id === undefined ? inCell.formula : state.formulas[id]!.text
      if (text !== undefined) out.formula = text
      if (inCell.format !== undefined) out.format = inCell.format
      if (inCell.meta !== undefined) out.meta = inCell.meta
      cells[addr] = out
      vals.set(addr, value)
    }
    values.set(sheet.name, vals)
    return { name: sheet.name, cells }
  })
  const volatile = [...state.volatile]
    .map((id) => state.formulas[id]!)
    .sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
    .map((f) => ({ sheet: state.sheetNames[f.s], cell: formatAddress(f.r, f.c) }))
  return {
    sheets,
    issues: issues(state).map((i) => ({ sheet: state.sheetNames[i.s], cell: formatAddress(i.r, i.c), error: i.error, reason: i.reason, ...(i.check ? { check: i.check, at: i.at } : {}) })),
    nameIssues: state.nameIssues,
    volatile,
    clock: { ...clock },
    hash: valuesHash(values, CONFIG_VERSION),
    inputHash: inputHash(state.input),
    configVersion: CONFIG_VERSION,
    ms: performance.now() - t0,
  }
}
