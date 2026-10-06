/** The computation result per the contract: values by sheet (spilled cells too), error reasons, volatile cells, hashes. */
import { inputHash, outNumber, valuesHash, type FillLine } from './canon.js'
import { CONFIG_VERSION } from './config.js'
import type { CellResult, CellValue, Clock, ComputeResult } from './contract.js'
import { key, toCellValue, valueAt, type EngineState } from './engine.js'
import { issues } from './issues.js'
import { nowMs } from './env.js'
import { formatAddress, msToSerial, parseAddress, sortAddresses } from './model.js'
import { ownerOf } from './spill/place.js'
import type { Spill } from './spill/state.js'


/** The clock moment as a serial number (UTC); an invalid moment is a caller error. */
export function clockSerial(clock: Clock): number {
  const ms = Date.parse(clock.now)
  if (Number.isNaN(ms)) throw new Error(`bad clock.now: ${clock.now}`)
  return msToSerial(ms)
}

export function outValue(v: CellValue): CellValue {
  return typeof v === 'number' ? outNumber(v) : v
}

const rangeText = (x: { r1: number; c1: number; r2: number; c2: number }) => `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`

/** A spilled formula's range and the value of its unstored cells. */
const anchorFields = (x: Spill): Pick<CellResult, 'spill' | 'spillFill'> => (x.fills.length ? { spill: rangeText(x), spillFill: outValue(toCellValue(x.fills[0].v)) } : { spill: rangeText(x) })

/** A formula's spill (`spill`, `spillFill`) or a spilled cell's formula (`spillOf`), as results and changes show them. */
export function spillFields(state: EngineState, s: number, r: number, c: number): Pick<CellResult, 'spill' | 'spillFill' | 'spillOf'> {
  const id = state.at.get(key(s, r, c))
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (x) return anchorFields(x)
  const o = ownerOf(state, s, r, c)
  return o ? { spillOf: formatAddress(o.r1, o.c1) } : {}
}

/** The stored cells of a sheet's spills, by address, with their formula's address. */
function spilled(state: EngineState, s: number): Map<string, string> {
  const out = new Map<string, string>()
  const xs: Spill[] = [...state.spills.on.values()].filter((x) => x.s === s)
  for (const x of xs) {
    const from = formatAddress(x.r1, x.c1)
    for (let r = x.r1; r < x.r1 + x.h; r++) for (let c = x.c1; c < x.c1 + x.w; c++) if ((r !== x.r1 || c !== x.c1) && !state.spills.occupied.has(s, r, c)) out.set(formatAddress(r, c), from)
  }
  return out
}

export function buildResult(state: EngineState, clock: Clock, t0: number): ComputeResult {
  const values = new Map<string, Map<string, CellValue>>()
  const fills: FillLine[] = []
  const sheets = state.input.sheets.map((sheet, s) => {
    const cells: Record<string, CellResult> = {}
    const vals = new Map<string, CellValue>()
    const spill = spilled(state, s)
    for (const addr of sortAddresses([...Object.keys(sheet.cells), ...spill.keys()])) {
      const { row, col } = parseAddress(addr)
      const inCell = sheet.cells[addr] ?? {}
      const value = outValue(toCellValue(valueAt(state, s, row, col)))
      const id = state.at.get(key(s, row, col))
      const out: CellResult = { value }
      const text = id === undefined ? inCell.formula : state.formulas[id]!.text
      if (text !== undefined) out.formula = text
      if (inCell.format !== undefined) out.format = inCell.format
      if (inCell.meta !== undefined) out.meta = inCell.meta
      const x = id === undefined ? undefined : state.spills.on.get(id)
      if (x) {
        Object.assign(out, anchorFields(x))
        for (const f of x.fills) fills.push({ sheet: sheet.name, range: rangeText(f), value: out.spillFill! })
      }
      const from = spill.get(addr)
      if (from) out.spillOf = from
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
    hash: valuesHash(values, CONFIG_VERSION, fills, state.input.locale),
    inputHash: inputHash(state.input),
    configVersion: CONFIG_VERSION,
    ms: nowMs() - t0,
  }
}
