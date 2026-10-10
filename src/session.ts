/**
 * A session over a book: edits are applied atomically with one recomputation of dependents. A live
 * session (`live`) takes a fresh clock and a new seed on every recomputation; its current clock
 * (`clock`) is what a saved result records, and a full computation with it gives the same hash.
 *
 * The operations are steps (`openSteps`): `openSession` runs each to its end at once, the asynchronous session in
 * portions. A call that stops (a limit, a deadline, an abort) changes nothing: what it changed in place is put back,
 * and the state is built again from the book on the next call.
 */
import type { BookInput, BookLayout, CellChange, CellEdit, CellInput, CellPatch, CellRef, CellStyle, CellValidation, CellView, Clock, ComputeResult, EntryCheck, Explanation, ExplainOptions, LayoutEdit, NameInput, NumberFormat, Scalar, Session, SessionHistory, SessionOptions, SessionRestructureResult, SheetExtent, SheetInput, StructureOp, UndoResult } from './contract.js'
import { applyClock, applyEdits, applyHidden, applyMerges, buildState, setNow, type EngineState } from './engine.js'
import { limitsBudget, run, untimed, withBudget, type Steps } from './exec/budget.js'
import { explainCell, explainText } from './explain.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkCell, checkClock, checkEdits, checkLayoutEdits, checkNames, checkSessionOptions, checkSheet, checkString, checkStyles, InputError } from './input.js'
import { applyLayout, bookLayout, checkStylesUsed, contentEdits, inputsOf } from './session/book.js'
import { checkArrayEdits } from './spill/fixed.js'
import { sheetExtent } from './session/extent.js'
import { cellValue, editedCells, layoutCells, shownChanges, snapshot } from './session/looks.js'
import { movedChanges } from './session/moved.js'
import { applyPatches, checkPatches } from './session/patch.js'
import { bookDiff, type BookDiff } from './session/rebook.js'
import { applyStep, bookStep, cellsStep, History, layoutStep, stylesStep, type HistoryData, type SessionCore } from './session/undo.js'
import { headerEdits } from './tables/headers.js'
import { checkTables } from './tables/check.js'
import { dropStats } from './cond/stats.js'
import { fitEdits } from './layout/fit.js'
import { hiddenRows } from './layout/hidden.js'
import { parseMerge } from './layout/merges.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, cellsIn, clockSerial, spillFields, stateHash } from './result.js'
import { cellLooks } from './shown.js'
import { parseRange } from './structure/area.js'
import { applyOp, checkedMapCell } from './structure/index.js'
import { stateValues } from './structure/values.js'
import { sheetKey } from './syntax.js'
import { cellValidation, checkEntry } from './validation/session.js'

/** How a session is run by its host. */
export interface SessionHooks {
  /** The clock a live session takes at each recomputation (a thread replaying a journal gives the recorded ones). */
  clock?: () => Clock
  /** The history a session restored from a snapshot starts with. */
  history?: HistoryData
}

/** A session's operations as steps; `book`, `setStyles` and `close` compute nothing. */
export interface SessionSteps {
  readonly clock: Clock
  setCells(edits: CellEdit[]): Steps<CellChange[]>
  patchCells(patches: CellPatch[]): Steps<CellChange[]>
  setNames(names: NameInput[]): Steps<CellChange[]>
  setClock(clock: Clock): Steps<CellChange[]>
  restructure(op: StructureOp): Steps<SessionRestructureResult>
  setStyles(styles: CellStyle[]): void
  setLayout(edits: LayoutEdit[]): Steps<CellChange[]>
  book(): BookInput
  layout(): BookLayout
  extent(sheet: string): Steps<SheetExtent>
  undo(): Steps<UndoResult | null>
  redo(): Steps<UndoResult | null>
  readonly history: SessionHistory
  /** What the history holds, for a snapshot. */
  historyData(): HistoryData
  explain(sheet: string, cell: string, opts?: ExplainOptions): Steps<Explanation>
  cells(sheet: string, range: string | string[]): Steps<Record<string, CellView>>
  validation(sheet: string, cell: string): Steps<CellValidation | null>
  validateEntry(sheet: string, cell: string, input: CellInput | null): Steps<EntryCheck>
  result(): Steps<ComputeResult>
  hash(): Steps<string>
  close(): void
}

/**
 * The name a number or a boolean typed into a header cell gives its column: as the cell's own number format shows it,
 * without the column's width (`1,50` by `0.00`, `31.01.2025` by the short date in ru-RU), or as `&` writes it where the
 * format shows it as General does (`0,000000001`, not `1E-09`): the set `tables/headers-ru-held`.
 */
function headerText(s: EngineState, v: Scalar, format: NumberFormat | undefined): string {
  const plain = String(s.ctx.rules.toText(v))
  if (format === undefined || typeof v !== 'number') return plain
  const shown = s.ctx.rules.display(v, format, Infinity).text
  return shown === s.ctx.rules.display(v, undefined, Infinity).text ? plain : shown
}

/** Opens a session in steps: the book checked, then computed. */
export function* openSteps(book: BookInput, opts?: SessionOptions, hooks: SessionHooks = {}): Steps<SessionSteps> {
  const o = checkSessionOptions(opts)
  checkBook(book)
  const live = o.live === true
  const fresh = hooks.clock ?? currentClock
  let clock: Clock = 'clock' in o && o.clock ? { ...o.clock } : fresh()
  const { externals, execution } = o
  const display = o.display !== false
  const build = (b: BookInput) => buildState(b, { now: clockSerial(clock), seed: clock.seed, externals, execution })
  let state: EngineState | null = yield* build(book)
  // a call stopped while changing the state in place: the state is built again from its book before the next one
  let stale = false
  const st = () => {
    if (!state) throw new Error('session is closed')
    return state
  }
  function* ready(): Steps<EngineState> {
    const s = st()
    if (stale) {
      // what an earlier call left to do does not count in this call's soft deadline
      state = yield* untimed(build(s.input))
      stale = false
    }
    return state!
  }
  const sheetIndex = (s: EngineState) => (name: string) => s.sheetIdx.get(sheetKey(name))
  const styleCount = (s: EngineState) => s.input.styles?.length ?? 0
  const changes = (s: EngineState, keys: string[]): CellChange[] => {
    // the statistics of conditional formats over changed cells are taken again
    dropStats(s, keys)
    const look = display ? cellLooks(s) : null
    return keys.map((k) => {
      const [i, r, c] = k.split(':').map(Number)
      const value = cellValue(s, i, r, c)
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value, ...look?.(i, r, c, value), ...spillFields(s, i, r, c) }
    })
  }
  // keys of changed values as the engine orders them, then the other cells whose look changed, by address
  const merged = (keys: string[], more: string[]) => {
    const seen = new Set(keys)
    const extra = more.filter((k) => !seen.has(k)).map((k) => k.split(':').map(Number))
    return [...keys, ...extra.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).map((k) => k.join(':'))]
  }
  const tick = (s: EngineState) => {
    clock = fresh()
    setNow(s, clockSerial(clock), clock.seed)
  }
  // a call that builds a new state aside: when it stops, the old state gets its clock back
  const untick = (s: EngineState, at: Clock) => {
    if (clock === at) return
    clock = at
    setNow(s, clockSerial(at), at.seed)
  }
  const history = new History(o.undo, hooks.history)
  // a step is made only when the history keeps steps (a book's diff costs a pass over its cells)
  const keep = (step: () => Parameters<History['push']>[0]) => history.on && history.push(step())
  /** Edits of cells (checked) applied; the inputs they replaced, which are put back if the call stops. */
  function* putCells(s: EngineState, edits: CellEdit[]): Steps<[CellChange[], CellEdit[]]> {
    const before = display ? snapshot(s, editedCells(s, edits)) : null
    const was = inputsOf(s, edits)
    const at = clock
    try {
      // edits that keep a cell's value and formula change its input alone
      const loud = contentEdits(s, edits)
      if (loud.length && live) tick(s)
      const keys = loud.length ? yield* applyEdits(s, loud, live) : []
      return [changes(s, before ? merged(keys, shownChanges(s, before)) : keys), was]
    } catch (e) {
      for (const { sheet, cell, input } of was) {
        const cells = s.input.sheets[sheetIndex(s)(sheet)!].cells
        if (input === null) delete cells[cell]
        else cells[cell] = input
      }
      clock = at
      stale = true
      throw e
    }
  }
  /**
   * The state built again from `book`, the session's book changed, with its changes against the old one by `mapCell`
   * (and the looks of both with `looks`) and the way back from the book before (`back`); nothing changes if it stops.
   */
  function* rebuild(old: EngineState, book: BookInput, mapCell: (ref: CellRef) => CellRef | null, looks: boolean, back: () => Steps<BookDiff | null>): Steps<[EngineState, CellChange[], BookDiff | null]> {
    const at = clock
    if (live) tick(old)
    try {
      const next = yield* build(book)
      const keys = movedChanges(old, next, mapCell, display && looks ? [cellLooks(old), cellLooks(next)] : undefined)
      const diff = yield* back()
      state = next
      return [next, changes(next, keys), diff]
    } catch (e) {
      untick(old, at)
      throw e
    }
  }
  /** The way back to `before` from `book` (given to `rebuild`), when the history keeps steps. */
  const diffTo = (book: BookInput, before: BookInput, op: StructureOp | null) =>
    function* (): Steps<BookDiff | null> {
      return history.on ? yield* bookDiff(book, before, op, 'back') : null
    }
  /**
   * Sheets of the book replaced by new ones of another layout: the arrays and SUBTOTALs it concerns computed again; the
   * changes and the sheets before, which are put back if the call stops.
   */
  function* relayout(s: EngineState, next: Map<number, SheetInput>, shown: Map<string, string> | null): Steps<[CellChange[], SheetInput[]]> {
    const sheets = s.input.sheets.slice()
    for (const [i, x] of next) s.input.sheets[i] = x
    if ([...next].some(([i, x]) => x.conditionalFormats !== sheets[i].conditionalFormats)) dropStats(s, [], true)
    try {
      // merged areas added or removed: the arrays whose range they meet spill or stop; hidden rows: SUBTOTAL and AGGREGATE count again
      const keys = new Set<string>()
      for (const [i, x] of next) {
        const [was, now] = [new Set(sheets[i].merges ?? []), new Set(x.merges ?? [])]
        const moved = [...was].filter((m) => !now.has(m)).concat([...now].filter((m) => !was.has(m)))
        if (moved.length) for (const k of yield* applyMerges(s, i, moved.map((m) => parseMerge(m)!))) keys.add(k)
        if (JSON.stringify(hiddenRows(x) ?? null) !== JSON.stringify(hiddenRows(sheets[i]) ?? null)) for (const k of yield* applyHidden(s, i)) keys.add(k)
      }
      return [changes(s, shown ? merged([...keys], shownChanges(s, shown)) : [...keys]), sheets]
    } catch (e) {
      s.input.sheets.splice(0, sheets.length, ...sheets)
      stale = true
      throw e
    }
  }
  const core: SessionCore = { ready, putCells, rebuild, relayout, display, sheetIndex }
  return {
    get clock() {
      return { ...clock }
    },
    *setCells(edits: CellEdit[]) {
      const s = yield* ready()
      checkEdits(edits, sheetIndex(s), styleCount(s))
      checkArrayEdits(s, edits)
      // a table's header cell typed over renames its column and the references to it: the book is built again
      const renamed = headerEdits(s.input, edits, (v, format) => headerText(s, v, format))
      if (renamed) {
        const [, out, diff] = yield* rebuild(s, renamed, (ref) => ref, true, diffTo(renamed, s.input, null))
        keep(() => diff && bookStep('setCells', diff, renamed))
        return out
      }
      const [out, was] = yield* putCells(s, edits)
      if (was.length) keep(() => cellsStep('setCells', was))
      return out
    },
    *patchCells(patches: CellPatch[]) {
      const s = yield* ready()
      const checked = checkPatches(patches, sheetIndex(s), styleCount(s))
      const before = display ? snapshot(s, editedCells(s, checked)) : null
      const was = inputsOf(s, checked)
      applyPatches(s, checked, (name) => sheetIndex(s)(name)!)
      if (was.length) keep(() => cellsStep('patchCells', was))
      return before ? changes(s, merged([], shownChanges(s, before))) : []
    },
    *setNames(names: NameInput[]) {
      const old = yield* ready()
      checkNames(names, 'names')
      // tables and names share one namespace
      checkTables({ ...old.input, names }, 'book')
      // names change the whole graph: the book is rebuilt, changes are collected over all cells (none moved)
      const book = { ...old.input, names }
      const [, out, diff] = yield* rebuild(old, book, (ref) => ref, false, diffTo(book, old.input, null))
      keep(() => diff && bookStep('setNames', diff, book))
      return out
    },
    *setClock(next: Clock) {
      if (live) throw new Error('a live session takes its clock from the system')
      const s = yield* ready()
      const at = clock
      clock = { ...checkClock(next, 'clock') }
      try {
        return changes(s, yield* applyClock(s, clockSerial(clock), clock.seed))
      } catch (e) {
        clock = at
        stale = true
        throw e
      }
    },
    *restructure(op: StructureOp) {
      const old = yield* ready()
      const r = applyOp(old.input, op, stateValues(old))
      if (!r.ok) return r
      // a sheet hidden or shown changes no value: the book takes its sheets, nothing is built again
      if (op.op === 'hideSheet' || op.op === 'unhideSheet') {
        const next = new Map(r.book.sheets.flatMap((x, i): [number, SheetInput][] => (x === old.input.sheets[i] ? [] : [[i, x]])))
        const [, before] = yield* relayout(old, next, null)
        keep(() => layoutStep('restructure', before, next))
        return { ok: true, book: old.input, mapCell: checkedMapCell(old.input, r.mapCell), changes: [] }
      }
      // the operation as it was given, kept for its geometry: the caller's object may change
      const kept = JSON.parse(JSON.stringify(op)) as StructureOp
      // the simple way: the state is built again from the new book
      const [next, out, diff] = yield* rebuild(old, r.book, r.mapCell, true, diffTo(r.book, old.input, kept))
      keep(() => diff && bookStep('restructure', diff, r.book))
      return { ok: true, book: next.input, mapCell: checkedMapCell(old.input, r.mapCell), changes: out }
    },
    setStyles(styles: CellStyle[]) {
      const s = st()
      checkStylesUsed(s.input, checkStyles(styles, 'styles').length)
      const was = s.input.styles
      s.input.styles = styles
      keep(() => stylesStep(was))
    },
    *setLayout(edits: LayoutEdit[]) {
      const s = yield* ready()
      // AutoFit becomes the widths it sets
      const checked = fitEdits(s, checkLayoutEdits(edits, sheetIndex(s), styleCount(s)))
      const shown = display ? snapshot(s, layoutCells(s, checked)) : null
      const next = applyLayout(s.input, checked, (name) => sheetIndex(s)(name)!)
      const [out, before] = yield* relayout(s, next, shown)
      if (edits.length) keep(() => layoutStep('setLayout', before, next))
      return out
    },
    book() {
      return st().input
    },
    layout() {
      return bookLayout(st().input)
    },
    *extent(sheet: string) {
      const s = yield* ready()
      return sheetExtent(s, checkSheet(sheet, sheetIndex(s), 'sheet'))
    },
    *explain(sheet: string, cell: string, o: ExplainOptions = {}) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const { row, col } = parseAddress(checkAddress(cell, 'cell'))
      const e = explainCell(s, i, row, col, o.depth ?? 3, o.maxNodes ?? 200, o.expand ?? false)
      return { ...e, text: explainText(e.root, s.input.locale) }
    },
    *cells(sheet: string, range: string | string[]) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const many = Array.isArray(range)
      const rects = (many ? range : [range]).map((text, k) => {
        const path = many ? `range[${k}]` : 'range'
        const x = parseRange(checkString(text, path))
        if (!x) throw new InputError(path, `expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got ${JSON.stringify(text)}`)
        return x
      })
      return yield* cellsIn(s, i, rects, display)
    },
    *validation(sheet: string, cell: string) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const { row, col } = parseAddress(checkAddress(cell, 'cell'))
      return cellValidation(s, i, row, col)
    },
    *validateEntry(sheet: string, cell: string, input: CellInput | null) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const { row, col } = parseAddress(checkAddress(cell, 'cell'))
      if (input !== null) checkCell(input, 'input', styleCount(s))
      checkArrayEdits(s, [{ sheet, cell, input }])
      try {
        return yield* checkEntry(s, i, row, col, input)
      } catch (e) {
        // the input put in for the check may still be there: the state is built again on the next call
        stale = true
        throw e
      }
    },
    *hash() {
      return yield* stateHash(yield* ready())
    },
    *result() {
      const t0 = nowMs()
      return yield* buildResult(yield* ready(), clock, t0, display)
    },
    *undo() {
      return yield* applyStep(core, history.undo.at(-1), (forth) => history.undone(forth))
    },
    *redo() {
      return yield* applyStep(core, history.redo.at(-1), (back) => history.redone(back))
    },
    get history() {
      return history.state
    },
    historyData: () => history.data(),
    close() {
      state = null
    },
  }
}

/**
 * Opens a session. Without options it keeps the clock of the moment it was opened (with a random seed)
 * until `setClock`; `{ live: true }` takes a fresh clock on every recomputation.
 */
export function openSession(book: BookInput, opts?: SessionOptions): Session {
  const budget = limitsBudget(checkSessionOptions(opts).limits)
  const now = <T>(steps: () => Steps<T>): T => withBudget(budget, () => run(steps()))
  const core = now(() => openSteps(book, opts))
  return {
    get clock() {
      return core.clock
    },
    setCells: (edits) => now(() => core.setCells(edits)),
    patchCells: (patches) => now(() => core.patchCells(patches)),
    setNames: (names) => now(() => core.setNames(names)),
    setClock: (clock) => now(() => core.setClock(clock)),
    restructure: (op) => now(() => core.restructure(op)),
    setStyles: (styles) => core.setStyles(styles),
    setLayout: (edits) => now(() => core.setLayout(edits)),
    book: () => core.book(),
    layout: () => core.layout(),
    extent: (sheet) => now(() => core.extent(sheet)),
    explain: (sheet, cell, o) => now(() => core.explain(sheet, cell, o)),
    cells: (sheet, range) => now(() => core.cells(sheet, range)),
    validation: (sheet, cell) => now(() => core.validation(sheet, cell)),
    validateEntry: (sheet, cell, input) => now(() => core.validateEntry(sheet, cell, input)),
    result: () => now(() => core.result()),
    hash: () => now(() => core.hash()),
    undo: () => now(() => core.undo()),
    redo: () => now(() => core.redo()),
    get history() {
      return core.history
    },
    close: () => core.close(),
  }
}
