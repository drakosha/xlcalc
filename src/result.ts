/** The computation result per the contract: values by sheet (spilled cells too), error reasons, volatile cells, hashes. */
import { inputHash, outNumber, valuesHash, type FillLine } from './canon.js'
import { CONFIG_VERSION } from './config.js'
import type { CellResult, CellValue, CellView, Clock, ComputeResult } from './contract.js'
import { key, toCellValue, valueAt, type EngineState } from './engine.js'
import { spend } from './exec/budget.js'
import { issues } from './issues.js'
import { nowMs } from './env.js'
import { formatAddress, msToSerial, parseAddress, sortAddresses } from './model.js'
import { cellLooks, type Looks } from './shown.js'
import { ownerOf } from './spill/place.js'
import type { Spill } from './spill/state.js'
import { meets, type Rect } from './structure/area.js'


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

/** The stored cells of a sheet's spills (those in `within`), by address, with their formula's address. */
function spilled(state: EngineState, s: number, within?: Rect): Map<string, string> {
  const out = new Map<string, string>()
  const xs: Spill[] = [...state.spills.on.values()].filter((x) => x.s === s && (!within || meets(x, within)))
  for (const x of xs) {
    const from = formatAddress(x.r1, x.c1)
    const [r1, r2] = within ? [Math.max(x.r1, within.r1), Math.min(x.r1 + x.h - 1, within.r2)] : [x.r1, x.r1 + x.h - 1]
    const [c1, c2] = within ? [Math.max(x.c1, within.c1), Math.min(x.c1 + x.w - 1, within.c2)] : [x.c1, x.c1 + x.w - 1]
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if ((r !== x.r1 || c !== x.c1) && !state.spills.occupied.has(s, r, c)) out.set(formatAddress(r, c), from)
  }
  return out
}

/** A cell of sheet `s` as a result gives it; `spillOf`, the formula of a spilled cell. */
function cellResult(state: EngineState, s: number, addr: string, look: Looks | null, spillOf: string | undefined): CellResult {
  const { row, col } = parseAddress(addr)
  const inCell = state.input.sheets[s].cells[addr] ?? {}
  const value = outValue(toCellValue(valueAt(state, s, row, col)))
  const id = state.at.get(key(s, row, col))
  const out: CellResult = look ? { value, ...look(s, row, col, value, inCell) } : { value }
  const text = id === undefined ? inCell.formula : state.formulas[id]!.text
  if (text !== undefined) out.formula = text
  if (inCell.format !== undefined) out.format = inCell.format
  if (inCell.meta !== undefined) out.meta = inCell.meta
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (x) Object.assign(out, anchorFields(x))
  if (spillOf) out.spillOf = spillOf
  return out
}

/** Rectangles up to this many cells are read cell by cell; a larger one by the sheet's cells. */
const BY_CELL = 65_536

/**
 * The cells of rectangle `x` of sheet `s` as the result gives them (`Session.cells`), with each cell's own style: the
 * rectangle's cells looked up one by one, or a larger one's found among the sheet's.
 */
export function cellsIn(state: EngineState, s: number, x: Rect, display: boolean): Record<string, CellView> {
  const cells = state.input.sheets[s].cells
  const look = display ? cellLooks(state) : null
  const spill = spilled(state, s, x)
  const inputs: string[] = []
  if ((x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1) <= BY_CELL) {
    for (let r = x.r1; r <= x.r2; r++)
      for (let c = x.c1; c <= x.c2; c++) {
        const a = formatAddress(r, c)
        if (a in cells) inputs.push(a)
      }
  } else
    for (const a of Object.keys(cells)) {
      const { row, col } = parseAddress(a)
      if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2) inputs.push(a)
    }
  const out: Record<string, CellView> = {}
  for (const addr of sortAddresses([...inputs, ...spill.keys()])) {
    spend(1)
    const view: CellView = cellResult(state, s, addr, look, spill.get(addr))
    const style = cells[addr]?.style
    if (style !== undefined) view.style = style
    out[addr] = view
  }
  return out
}

/** The result; with `display` each cell's shown text and colour after its value. */
export function buildResult(state: EngineState, clock: Clock, t0: number, display: boolean): ComputeResult {
  const values = new Map<string, Map<string, CellValue>>()
  const fills: FillLine[] = []
  const look = display ? cellLooks(state) : null
  const sheets = state.input.sheets.map((sheet, s) => {
    const cells: Record<string, CellResult> = {}
    const vals = new Map<string, CellValue>()
    const spill = spilled(state, s)
    for (const addr of sortAddresses([...Object.keys(sheet.cells), ...spill.keys()])) {
      spend(1)
      const out = cellResult(state, s, addr, look, spill.get(addr))
      if (out.spillFill !== undefined) {
        const { row, col } = parseAddress(addr)
        for (const f of state.spills.on.get(state.at.get(key(s, row, col))!)!.fills) fills.push({ sheet: sheet.name, range: rangeText(f), value: out.spillFill })
      }
      cells[addr] = out
      vals.set(addr, out.value)
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
    hash: valuesHash(values, CONFIG_VERSION, fills, state.input.locale, state.input.dateSystem, true),
    inputHash: inputHash(state.input),
    configVersion: CONFIG_VERSION,
    ms: nowMs() - t0,
  }
}
