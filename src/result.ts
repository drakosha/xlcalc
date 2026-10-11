/** The computation result per the contract: values by sheet (spilled cells too), error reasons, volatile cells, hashes. */
import { inputHashSteps, outNumber, sheetOrder, sortedSteps, ValuesHasher, type FillLine } from './canon.js'
import { CONFIG_VERSION } from './config.js'
import type { CellResult, CellValue, CellView, Clock, ComputeResult, ConditionalAppearance, Diagnostic } from './contract.js'
import { key, toCellValue, valueAt, type EngineState } from './engine.js'
import { paced, type Steps } from './exec/budget.js'
import { issues } from './issues.js'
import { nowMs } from './env.js'
import { formatAddress, MAX_COLS, msToSerial, parseAddress } from './model.js'
import { windowLooks, withLook } from './cond/looks.js'
import { cellLooks, type Looks } from './shown.js'
import { ownerOf } from './spill/place.js'
import type { Spill } from './spill/state.js'
import { cellCount, cornersText, meets, type Rect } from './structure/area.js'


/** The clock moment as a serial number (UTC); an invalid moment is a caller error. */
export function clockSerial(clock: Clock): number {
  const ms = Date.parse(clock.now)
  if (Number.isNaN(ms)) throw new Error(`bad clock.now: ${clock.now}`)
  return msToSerial(ms)
}

export function outValue(v: CellValue): CellValue {
  return typeof v === 'number' ? outNumber(v) : v
}

/** A spilled formula's range and the value of its unstored cells. */
const anchorFields = (x: Spill): Pick<CellResult, 'spill' | 'spillFill'> => (x.fills.length ? { spill: cornersText(x), spillFill: outValue(toCellValue(x.fills[0].v)) } : { spill: cornersText(x) })

/** A formula's spill (`spill`, `spillFill`) or a spilled cell's formula (`spillOf`), as results and changes show them. */
export function spillFields(state: EngineState, s: number, r: number, c: number): Pick<CellResult, 'spill' | 'spillFill' | 'spillOf'> {
  const id = state.at.get(key(s, r, c))
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (x) return anchorFields(x)
  const o = ownerOf(state, s, r, c)
  return o ? { spillOf: formatAddress(o.r1, o.c1) } : {}
}

/** The stored cells of a sheet's spills (those in `within`), by address, with their formula's address; `whole`: a large spill's unstored rest too. */
export function* spilled(state: EngineState, s: number, within?: Rect, whole = false): Steps<Map<string, string>> {
  const out = new Map<string, string>()
  const xs: Spill[] = [...state.spills.on.values()].filter((x) => x.s === s && (!within || meets(x, within)))
  for (const x of xs) {
    const from = formatAddress(x.r1, x.c1)
    const [last, right] = whole ? [x.r2, x.c2] : [x.r1 + x.h - 1, x.c1 + x.w - 1]
    const [r1, r2] = within ? [Math.max(x.r1, within.r1), Math.min(last, within.r2)] : [x.r1, last]
    const [c1, c2] = within ? [Math.max(x.c1, within.c1), Math.min(right, within.c2)] : [x.c1, right]
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) {
        if (paced(1)) yield
        if ((r !== x.r1 || c !== x.c1) && !state.spills.occupied.has(s, r, c)) out.set(formatAddress(r, c), from)
      }
  }
  return out
}

/** A cell of sheet `s` as a result gives it; `spillOf`, the formula of a spilled cell; `format`, with the cell's own number format. */
export function cellResult(state: EngineState, s: number, addr: string, look: Looks | null, spillOf: string | undefined, format = true): CellResult {
  const { row, col } = parseAddress(addr)
  const inCell = state.input.sheets[s].cells[addr] ?? {}
  const value = outValue(toCellValue(valueAt(state, s, row, col)))
  const id = state.at.get(key(s, row, col))
  const out: CellResult = look ? { value, ...look(s, row, col, value, inCell) } : { value }
  const text = id === undefined ? inCell.formula : state.formulas[id]!.text
  if (text !== undefined) out.formula = text
  if (format && inCell.numberFormat !== undefined) out.numberFormat = inCell.numberFormat
  if (inCell.meta !== undefined) out.meta = inCell.meta
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (x) Object.assign(out, anchorFields(x))
  if (spillOf) out.spillOf = spillOf
  return out
}

/** Rectangles up to this many cells are read cell by cell; a larger one by the sheet's cells. */
const BY_CELL = 65_536

/**
 * The cells of rectangles `xs` of sheet `s` as the result gives them (`Session.getCells`), with each cell's own style, once
 * each: a rectangle's cells looked up one by one, or a larger one's found among the sheet's. The steps yield between cells.
 */
export function* cellsIn(state: EngineState, s: number, xs: Rect[], display: boolean): Steps<Record<string, CellView>> {
  const cells = state.input.sheets[s].cells
  const look = display ? cellLooks(state) : null
  const spill = new Map<string, string>()
  const listed = new Set<string>()
  const cond = new Map<number, ConditionalAppearance>()
  for (const x of xs) {
    const small = cellCount(x) <= BY_CELL
    const mine: string[] = []
    for (const [a, from] of yield* spilled(state, s, x)) {
      spill.set(a, from)
      mine.push(a)
    }
    if (small) {
      for (let r = x.r1; r <= x.r2; r++)
        for (let c = x.c1; c <= x.c2; c++) {
          if (paced(1)) yield
          const a = formatAddress(r, c)
          if (a in cells) mine.push(a)
        }
    } else
      for (const a of Object.keys(cells)) {
        if (paced(1)) yield
        const { row, col } = parseAddress(a)
        if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2) mine.push(a)
      }
    for (const a of mine) listed.add(a)
    // what the sheet's conditional formats give the cells: every cell of a window, the listed ones of a larger rectangle
    for (const [k, v] of yield* windowLooks(state, s, x, small ? undefined : new Set(mine.map((a) => cellIndex(a))))) cond.set(k, v)
  }
  for (const k of cond.keys()) listed.add(formatAddress(Math.floor(k / MAX_COLS), k % MAX_COLS))
  const out: Record<string, CellView> = {}
  for (const addr of yield* sortedSteps([...listed])) {
    if (paced(1)) yield
    const view: CellView = cellResult(state, s, addr, look, spill.get(addr))
    const input = cells[addr]
    // what the cell shows beside its value: its own style, note and link, an old-kind array's area, and its conditional formats' appearance
    if (input?.style !== undefined) view.style = input.style
    if (input?.note !== undefined) view.note = input.note
    if (input?.link !== undefined) view.link = input.link
    if (input?.array !== undefined) view.array = input.array
    const cf = cond.get(cellIndex(addr))
    if (cf) withLook(view, cf, look && ((format) => look(s, ...rowCol(addr), view.value, { ...input, numberFormat: format })))
    out[addr] = view
  }
  return out
}

const cellIndex = (addr: string) => {
  const { row, col } = parseAddress(addr)
  return row * MAX_COLS + col
}

const rowCol = (addr: string): [number, number] => {
  const { row, col } = parseAddress(addr)
  return [row, col]
}

/**
 * The cells of sheet `s` a result lists, in order (input cells and the stored cells of spills), each with the formula it
 * spilled from; the steps yield between cells when the computation's slice is used up.
 */
function* eachCell(state: EngineState, s: number, f: (addr: string, spillOf: string | undefined) => void): Steps<void> {
  const spill = yield* spilled(state, s)
  let last: string | undefined
  for (const addr of yield* sortedSteps([...Object.keys(state.input.sheets[s].cells), ...spill.keys()])) {
    if (paced(1)) yield
    // once: an input cell that does not block a spill (a format or a style alone) is a spilled one too
    if (addr === last) continue
    last = addr
    f(addr, spill.get(addr))
  }
}

/** The hash's lines for the unstored cells of a large spill whose formula is at `addr` (none for other cells). */
function fillLines(state: EngineState, s: number, addr: string, value: CellValue | undefined, out: FillLine[]) {
  if (value === undefined) return
  const { row, col } = parseAddress(addr)
  for (const f of state.spills.on.get(state.at.get(key(s, row, col))!)!.fills) out.push({ sheet: state.sheetNames[s], range: cornersText(f), value })
}

/** The values hasher of a state's configuration, locale and date system; cells are written to it as they are read. */
const hasher = (state: EngineState) => new ValuesHasher(CONFIG_VERSION, state.input.locale, state.input.dateSystem)

/** The values hash of a state, as its result gives it, without the cells' results (`Session.getHash`). */
export function* stateHash(state: EngineState): Steps<string> {
  const w = hasher(state)
  const fills: FillLine[] = []
  for (const s of sheetOrder(state.input.sheets.map((x) => x.name))) {
    const quoted = JSON.stringify(state.input.sheets[s].name)
    yield* eachCell(state, s, (addr) => {
      const { row, col } = parseAddress(addr)
      w.cell(quoted, addr, outValue(toCellValue(valueAt(state, s, row, col))), true)
      const id = state.at.get(key(s, row, col))
      const x = id === undefined ? undefined : state.spills.on.get(id)
      if (x?.fills.length) fillLines(state, s, addr, anchorFields(x).spillFill, fills)
    })
  }
  return w.hex(fills)
}

/**
 * The result; with `display` each cell's shown text and colour after its value. Sheets are read in the hash's order,
 * listed in the book's. The steps yield between cells, so a large book's result does not hold the thread.
 */
export function* buildResult(state: EngineState, clock: Clock, t0: number, display: boolean): Steps<ComputeResult> {
  const w = hasher(state)
  const fills: FillLine[] = []
  const look = display ? cellLooks(state) : null
  const sheets = state.input.sheets.map(({ name }) => ({ name, cells: {} as Record<string, CellResult> }))
  for (const s of sheetOrder(sheets.map((x) => x.name))) {
    const { name, cells } = sheets[s]
    const quoted = JSON.stringify(name)
    yield* eachCell(state, s, (addr, spillOf) => {
      const out = cellResult(state, s, addr, look, spillOf)
      fillLines(state, s, addr, out.spillFill, fills)
      cells[addr] = out
      w.cell(quoted, addr, out.value, true)
    })
  }
  const inputHash = yield* inputHashSteps(state.input)
  const volatile = [...state.volatile]
    .map((id) => state.formulas[id]!)
    .sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
    .map((f) => ({ sheet: state.sheetNames[f.s], cell: formatAddress(f.r, f.c) }))
  return {
    sheets,
    diagnostics: [
      ...issues(state).map((i): Diagnostic => ({ kind: 'cell', sheet: state.sheetNames[i.s], cell: formatAddress(i.r, i.c), error: i.error, reason: i.reason, ...(i.check ? { check: i.check, at: i.at } : {}) })),
      ...state.nameIssues.map((n): Diagnostic => ({ kind: 'name', ...n })),
    ],
    volatile,
    clock: { ...clock },
    hash: w.hex(fills),
    inputHash,
    configVersion: CONFIG_VERSION,
    ms: nowMs() - t0,
  }
}
