/** The computation result per the contract: values by sheet (spilled cells too), error reasons, volatile cells, hashes. */
import { inputHashSteps, outNumber, sheetOrder, sortedSteps, ValuesHasher, type FillLine } from './canon.js'
import { CONFIG_VERSION } from './config.js'
import type { CellResult, CellValue, CellView, Clock, ComputeResult, ConditionalAppearance, Diagnostic, NumberFormat } from './contract.js'
import { toCellValue, valueAt, type EngineState } from './engine.js'
import { paced, type Steps } from './exec/budget.js'
import { issues } from './issues.js'
import { nowMs } from './env.js'
import { addressOrder, formatAddress, MAX_COLS, msToSerial, parseAddress, placeOf } from './model.js'
import type { V } from './values.js'
import { windowLooks, withLook } from './cond/looks.js'
import { formatCodeIn } from './locale/rules.js'
import { cellLooks, type Looks } from './shown.js'
import { NONE, type Layer } from './looks/layer.js'
import { inputIn, layerOf } from './looks/sheet.js'
import { ownerOf } from './spill/place.js'
import type { Spill } from './spill/state.js'
import { cellCount, cornersText, meets, type Rect } from './structure/area.js'


/** The clock moment as a serial number (UTC); an invalid moment is a caller error. */
export function clockSerial(clock: Clock): number {
  const ms = Date.parse(clock.now)
  if (Number.isNaN(ms)) throw new Error(`invalid clock.now: ${clock.now}`)
  return msToSerial(ms)
}

export function outValue(v: CellValue): CellValue {
  return typeof v === 'number' ? outNumber(v) : v
}

/** A value the engine holds as a result gives it: a number to 15 significant digits, an error as `{ error }`, null for none. */
export const outCell = (v: V): CellValue => outValue(toCellValue(v))

/** A cell's value as a result shows it. */
export const cellValue = (state: EngineState, s: number, r: number, c: number): CellValue => outCell(valueAt(state, s, r, c))

/** The engine options a clock gives: its moment as a serial number, and its seed. */
export const clockOptions = (clock: Clock) => ({ now: clockSerial(clock), seed: clock.seed })

/** A spilled formula's range and the value of its unstored cells. */
const anchorFields = (x: Spill): Pick<CellResult, 'spill' | 'spillFill'> => (x.fills.length ? { spill: cornersText(x), spillFill: outCell(x.fills[0].v) } : { spill: cornersText(x) })

/** A formula's spill (`spill`, `spillFill`) or a spilled cell's formula (`spillOf`), as results and changes show them. */
export function spillFields(state: EngineState, s: number, r: number, c: number): Pick<CellResult, 'spill' | 'spillFill' | 'spillOf'> {
  if (!state.spills.on.size) return {}
  const id = state.at.get(placeOf(s, r, c))
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

/**
 * A cell of sheet `s` as a result gives it; `spillOf`, the formula of a spilled cell; `format`, with the cell's own number
 * format (of its input, or of the sheet's looks for a cell without one).
 */
export function cellResult(state: EngineState, s: number, addr: string, look: Looks | null, spillOf: string | undefined, format = true): CellResult {
  const { row, col } = parseAddress(addr)
  return resultAt(state, s, addr, row, col, look, spillOf, format)
}

/** `cellResult` of the cell at `addr`, row `row` and column `col`: for callers that have both, without reading the address again. */
function resultAt(state: EngineState, s: number, addr: string, row: number, col: number, look: Looks | null, spillOf: string | undefined, format = true): CellResult {
  const inCell = inputIn(state.input.sheets[s], addr, row, col) ?? {}
  const value = cellValue(state, s, row, col)
  const id = state.at.get(placeOf(s, row, col))
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
 * each: a rectangle's cells looked up one by one, or a larger one's found among the sheet's; a cell of a look of the
 * sheet's looks too, as an input of a look alone would be. The steps yield between cells.
 */
export function* cellsIn(state: EngineState, s: number, xs: Rect[], display: boolean): Steps<Record<string, CellView>> {
  const sheet = state.input.sheets[s]
  const cells = sheet.cells
  const layer = layerOf(sheet)
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
          if (a in cells || (layer && layer.idAt(r, c) !== NONE)) mine.push(a)
        }
    } else {
      for (const a of Object.keys(cells)) {
        if (paced(1)) yield
        const { row, col } = parseAddress(a)
        if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2) mine.push(a)
      }
      for (const p of layer?.pieces(x) ?? [])
        for (let r = p.r1; r <= p.r2; r++)
          for (let c = p.c1; c <= p.c2; c++) {
            if (paced(1)) yield
            const a = formatAddress(r, c)
            if (!(a in cells)) mine.push(a)
          }
    }
    for (const a of mine) listed.add(a)
    // what the sheet's conditional formats give the cells: every cell of a window, the listed ones of a larger rectangle
    for (const [k, v] of yield* windowLooks(state, s, x, small ? undefined : new Set(mine.map(addressOrder)))) cond.set(k, v)
  }
  for (const k of cond.keys()) listed.add(formatAddress(Math.floor(k / MAX_COLS), k % MAX_COLS))
  const out: Record<string, CellView> = {}
  const code = formatCodeIn(state.input.locale)
  for (const addr of yield* sortedSteps([...listed])) {
    if (paced(1)) yield
    const { row, col } = parseAddress(addr)
    out[addr] = viewOf(state, s, addr, row, col, layer, look, spill.get(addr), cond.get(row * MAX_COLS + col), code)
  }
  return out
}

/**
 * A listed cell of `cellsIn` (its address `addr`, row `r` and column `c`): its result, and what it shows beside its value:
 * its own style, note and link, an old-kind array's area, its format's code, and its conditional formats' appearance `cf`.
 */
function viewOf(state: EngineState, s: number, addr: string, r: number, c: number, layer: Layer | undefined, look: Looks | null, spillOf: string | undefined, cf: ConditionalAppearance | undefined, code: (format: NumberFormat) => string): CellView {
  const view: CellView = resultAt(state, s, addr, r, c, look, spillOf)
  const input = state.input.sheets[s].cells[addr] ?? (layer && layer.lookAt(r, c))
  if (input?.style !== undefined) view.style = input.style
  // copies, as a pool's are: a host that edits them does not change the workbook
  if (input?.note !== undefined) view.note = { ...input.note }
  if (input?.link !== undefined) view.link = { ...input.link }
  if (input?.array !== undefined) view.array = input.array
  if (view.numberFormat !== undefined) view.numberFormatCode = code(view.numberFormat)
  if (cf) withLook(view, cf, look && ((format) => look(s, r, c, view.value, { ...input, numberFormat: format })))
  return view
}

/**
 * The cell at row `r` and column `c` of sheet `s` as `cellsIn` gives it for its one-cell rectangle (`Session.getCell`),
 * without the work of a rectangle; undefined where it lists none. `whole`: a cell of a large spill's unstored rest too,
 * as `spilled` takes it (`readCells`, which lists such cells). The steps yield as `spilled` does between a spill's
 * cells, and in the conditional formats' tests.
 */
export function* cellAt(state: EngineState, s: number, r: number, c: number, display: boolean, whole = false): Steps<CellView | undefined> {
  const cond = state.input.sheets[s].conditionalFormats?.length ? yield* windowLooks(state, s, { r1: r, c1: c, r2: r, c2: c }) : null
  return yield* viewAt(state, s, r, c, display ? cellLooks(state) : null, formatCodeIn(state.input.locale), cond, whole)
}

/**
 * `cellAt` (display on) of each cell of `cells` (sheet, row, column) in turn, the looks and each sheet's conditional
 * formats made once for them all (`readCells`, `findCells`): over the cells' own rectangles, not the box around them.
 */
export function* cellsAt(state: EngineState, cells: [s: number, r: number, c: number, ...rest: unknown[]][], whole: boolean): Steps<(CellView | undefined)[]> {
  const look = cellLooks(state)
  const code = formatCodeIn(state.input.locale)
  const conds = new Map<number, Map<number, ConditionalAppearance>>()
  for (const s of new Set(cells.map((x) => x[0])))
    if (state.input.sheets[s].conditionalFormats?.length) conds.set(s, yield* windowLooks(state, s, cells.filter((x) => x[0] === s).map(([, r, c]) => ({ r1: r, c1: c, r2: r, c2: c }))))
  const out: (CellView | undefined)[] = []
  for (const [s, r, c] of cells) out.push(yield* viewAt(state, s, r, c, look, code, conds.get(s) ?? null, whole))
  return out
}

/** `cellAt`'s view by looks `look`, codes `code` and the appearances `cond` of the sheet's conditional formats (null without any). */
function* viewAt(state: EngineState, s: number, r: number, c: number, look: Looks | null, code: (format: NumberFormat) => string, cond: Map<number, ConditionalAppearance> | null, whole: boolean): Steps<CellView | undefined> {
  const sheet = state.input.sheets[s]
  const addr = formatAddress(r, c)
  const layer = layerOf(sheet)
  const spillOf = state.spills.on.size ? (yield* spilled(state, s, { r1: r, c1: c, r2: r, c2: c }, whole)).get(addr) : undefined
  const k = r * MAX_COLS + c
  if (spillOf === undefined && !cond?.has(k) && !(addr in sheet.cells) && !(layer && layer.idAt(r, c) !== NONE)) return undefined
  return viewOf(state, s, addr, r, c, layer, look, spillOf, cond?.get(k), code)
}

/**
 * The cells of sheet `s` a result lists, in order (input cells and the stored cells of spills), each with the formula it
 * spilled from; the steps yield between cells when the computation's slice is used up. A cell of the sheet's looks is
 * listed where an input of a look alone was apart from the spills' stored cells: in the unstored part of a large spill.
 */
function* eachCell(state: EngineState, s: number, f: (addr: string, spillOf: string | undefined, row: number, col: number) => void, sorted?: Map<object, string[]>, keys?: Map<object, string[]>): Steps<void> {
  const spill = yield* spilled(state, s)
  const layer = state.spills.on.size ? layerOf(state.input.sheets[s]) : undefined
  const unstored: string[] = []
  if (layer)
    for (const x of state.spills.on.values())
      if (x.s === s)
        for (const fill of x.fills)
          for (const p of layer.pieces(fill))
            for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) unstored.push(formatAddress(r, c))
  const cells = state.input.sheets[s].cells
  const only = !spill.size && !unstored.length
  const places: number[] = []
  const listed = keys?.get(cells) ?? Object.keys(cells)
  const addrs = yield* sortedSteps(only ? listed : [...listed, ...spill.keys(), ...unstored], places)
  // the input cells alone: the input hash takes them in this order too
  if (only) sorted?.set(cells, addrs)
  let last: string | undefined
  for (let i = 0; i < addrs.length; i++) {
    if (paced(1)) yield
    const addr = addrs[i]
    // once: an input cell that does not block a spill (a format or a style alone) is a spilled one too
    if (addr === last) continue
    last = addr
    const col = places[i] % MAX_COLS
    f(addr, spill.get(addr), (places[i] - col) / MAX_COLS, col)
  }
}

/** The hash's lines for the unstored cells of a large spill whose formula is at `addr` (none for other cells). */
function fillLines(state: EngineState, s: number, addr: string, value: CellValue | undefined, out: FillLine[]) {
  if (value === undefined) return
  const { row, col } = parseAddress(addr)
  for (const f of state.spills.on.get(state.at.get(placeOf(s, row, col))!)!.fills) out.push({ sheet: state.sheetNames[s], range: cornersText(f), value })
}

/** The state's cells written to a hasher of its configuration, in the hash's sheet order, with each spill's fills: one walk for the values hash alone and a result. */
function* hashCells(state: EngineState, cell: (s: number, addr: string, spillOf: string | undefined, row: number, col: number) => { value: CellValue; spillFill?: CellValue }, sorted?: Map<object, string[]>, keys?: Map<object, string[]>): Steps<string> {
  const w = new ValuesHasher(CONFIG_VERSION, state.input.locale, state.input.dateSystem)
  const fills: FillLine[] = []
  for (const s of sheetOrder(state.input.sheets.map((x) => x.name))) {
    const quoted = JSON.stringify(state.input.sheets[s].name)
    yield* eachCell(state, s, (addr, spillOf, row, col) => {
      const { value, spillFill } = cell(s, addr, spillOf, row, col)
      w.cell(quoted, addr, value, true)
      if (spillFill !== undefined) fillLines(state, s, addr, spillFill, fills)
    }, sorted, keys)
  }
  return w.hex(fills)
}

/** The values hash of a state, as its result gives it, without the cells' results (`Session.getHash`). */
export function* stateHash(state: EngineState): Steps<string> {
  return yield* hashCells(state, (s, addr, _, row, col) => {
    const id = state.at.get(placeOf(s, row, col))
    const x = id === undefined ? undefined : state.spills.on.get(id)
    return { value: cellValue(state, s, row, col), spillFill: x?.fills.length ? anchorFields(x).spillFill : undefined }
  })
}

/**
 * The result; with `display` each cell's shown text and colour after its value. Sheets are read in the hash's order,
 * listed in the book's. The steps yield between cells, so a large book's result does not hold the thread.
 */
export function* buildResult(state: EngineState, clock: Clock, t0: number, display: boolean, book = state.input, keys?: Map<object, string[]>): Steps<ComputeResult> {
  const look = display ? cellLooks(state) : null
  const sheets = state.input.sheets.map(({ name }) => ({ name, cells: {} as Record<string, CellResult> }))
  const sorted = new Map<object, string[]>()
  const hash = yield* hashCells(state, (s, addr, spillOf, row, col) => {
    const out = resultAt(state, s, addr, row, col, look, spillOf)
    // A snapshot's result takes its metadata from the detached book, not from the session's live input.
    if (out.meta !== undefined && book !== state.input) out.meta = book.sheets[s].cells[addr].meta
    sheets[s].cells[addr] = out
    return { value: out.value, spillFill: out.spillFill }
  }, sorted, keys)
  const inputHash = yield* inputHashSteps(state.input, sorted)
  const volatile = [...state.volatile]
    .map((id) => state.formulas[id]!)
    .sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
    .map((f) => ({ sheet: state.sheetNames[f.s], cell: formatAddress(f.r, f.c) }))
  return {
    sheets,
    diagnostics: [
      ...issues(state).map((i): Diagnostic => ({ kind: 'cell', sheet: state.sheetNames[i.s], cell: formatAddress(i.r, i.c), error: i.error, reason: i.reason, ...(i.check ? { check: i.check, at: i.at } : {}), ...(i.fn !== undefined ? { fn: i.fn } : {}) })),
      ...state.nameIssues.map((n): Diagnostic => ({ kind: 'name', ...n })),
    ],
    volatile,
    clock: { ...clock },
    hash,
    inputHash,
    configVersion: CONFIG_VERSION,
    ms: nowMs() - t0,
  }
}
