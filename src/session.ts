/**
 * A session over a book: edits are applied atomically with one recomputation of dependents. A live
 * session (`live`) takes a fresh clock and a new seed on every recomputation; its current clock
 * (`clock`) is what a saved result records, and a full computation with it gives the same hash.
 *
 * The operations are steps (`openSteps`): `openSession` runs each to its end at once, the asynchronous session in
 * portions. A call that stops (a limit, a deadline, an abort) changes nothing: what it changed in place is put back,
 * and the state is built again from the book on the next call.
 */
import type { BookInput, CellChange, CellEdit, CellInput, CellStyle, Clock, ComputeResult, Explanation, ExplainOptions, LayoutEdit, NameInput, Session, SessionOptions, SessionRestructureResult, StructureOp } from './contract.js'
import { applyClock, applyEdits, applyHidden, applyMerges, buildState, setNow, type EngineState } from './engine.js'
import { limitsBudget, run, withBudget, type Steps } from './exec/budget.js'
import { explainCell, explainText } from './explain.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkClock, checkEdits, checkLayoutEdits, checkNames, checkSessionOptions, checkSheet, checkStyles } from './input.js'
import { applyLayout, checkStylesUsed, contentEdits } from './session/book.js'
import { checkArrayEdits } from './spill/fixed.js'
import { cellValue, editedCells, layoutCells, shownChanges, snapshot } from './session/looks.js'
import { movedChanges } from './session/moved.js'
import { headerEdits } from './tables/headers.js'
import { hiddenRows } from './layout/hidden.js'
import { parseMerge } from './layout/merges.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, clockSerial, spillFields } from './result.js'
import { cellLooks } from './shown.js'
import { applyOp, checkedMapCell } from './structure/index.js'
import { stateValues } from './structure/values.js'
import { sheetKey } from './syntax.js'

/** How a session is run by its host. */
export interface SessionHooks {
  /** The clock a live session takes at each recomputation (a thread replaying a journal gives the recorded ones). */
  clock?: () => Clock
}

/** A session's operations as steps; `book`, `setStyles` and `close` compute nothing. */
export interface SessionSteps {
  readonly clock: Clock
  setCells(edits: CellEdit[]): Steps<CellChange[]>
  setNames(names: NameInput[]): Steps<CellChange[]>
  setClock(clock: Clock): Steps<CellChange[]>
  restructure(op: StructureOp): Steps<SessionRestructureResult>
  setStyles(styles: CellStyle[]): void
  setLayout(edits: LayoutEdit[]): Steps<CellChange[]>
  book(): BookInput
  explain(sheet: string, cell: string, opts?: ExplainOptions): Steps<Explanation>
  result(): Steps<ComputeResult>
  close(): void
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
      state = yield* build(s.input)
      stale = false
    }
    return state!
  }
  const sheetIndex = (s: EngineState) => (name: string) => s.sheetIdx.get(sheetKey(name))
  const styleCount = (s: EngineState) => s.input.styles?.length ?? 0
  const changes = (s: EngineState, keys: string[]): CellChange[] => {
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
  return {
    get clock() {
      return { ...clock }
    },
    *setCells(edits: CellEdit[]) {
      const s = yield* ready()
      checkEdits(edits, sheetIndex(s), styleCount(s))
      checkArrayEdits(s, edits)
      // a table's header cell typed over renames its column and the references to it: the book is built again
      const renamed = headerEdits(s.input, edits, (v) => String(s.ctx.rules.toText(v)))
      if (renamed) {
        const at = clock
        if (live) tick(s)
        try {
          const next = yield* build(renamed)
          const keys = movedChanges(s, next, (ref) => ref, display ? [cellLooks(s), cellLooks(next)] : undefined)
          state = next
          return changes(next, keys)
        } catch (e) {
          untick(s, at)
          throw e
        }
      }
      const before = display ? snapshot(s, editedCells(s, edits)) : null
      // the inputs the edits replace, to put back if the call stops
      const was = new Map<string, [Record<string, CellInput>, string, CellInput | undefined]>()
      for (const e of edits) {
        const cells = s.input.sheets[sheetIndex(s)(e.sheet)!].cells
        const k = `${sheetKey(e.sheet)}!${e.cell}`
        if (!was.has(k)) was.set(k, [cells, e.cell, cells[e.cell]])
      }
      const at = clock
      try {
        // edits that keep a cell's value and formula change its input alone
        const loud = contentEdits(s, edits)
        if (loud.length && live) tick(s)
        const keys = loud.length ? yield* applyEdits(s, loud, live) : []
        return changes(s, before ? merged(keys, shownChanges(s, before)) : keys)
      } catch (e) {
        for (const [cells, addr, old] of was.values()) {
          if (old === undefined) delete cells[addr]
          else cells[addr] = old
        }
        clock = at
        stale = true
        throw e
      }
    },
    *setNames(names: NameInput[]) {
      const old = yield* ready()
      checkNames(names, 'names')
      const at = clock
      if (live) tick(old)
      try {
        // names change the whole graph: the book is rebuilt, changes are collected over all cells (none moved)
        const next = yield* build({ ...old.input, names })
        const keys = movedChanges(old, next, (ref) => ref)
        state = next
        return changes(next, keys)
      } catch (e) {
        untick(old, at)
        throw e
      }
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
      const at = clock
      if (live) tick(old)
      try {
        // the simple way: the state is built again from the new book
        const next = yield* build(r.book)
        const keys = movedChanges(old, next, r.mapCell, display ? [cellLooks(old), cellLooks(next)] : undefined)
        state = next
        return { ok: true, book: next.input, mapCell: checkedMapCell(old.input, r.mapCell), changes: changes(next, keys) }
      } catch (e) {
        untick(old, at)
        throw e
      }
    },
    setStyles(styles: CellStyle[]) {
      const s = st()
      checkStylesUsed(s.input, checkStyles(styles, 'styles').length)
      s.input.styles = styles
    },
    *setLayout(edits: LayoutEdit[]) {
      const s = yield* ready()
      const checked = checkLayoutEdits(edits, sheetIndex(s), styleCount(s))
      const shown = display ? snapshot(s, layoutCells(s, checked)) : null
      const before = s.input.sheets.map((x) => new Set(x.merges ?? []))
      const hidden = s.input.sheets.map((x) => JSON.stringify(hiddenRows(x) ?? null))
      // the sheets the layout replaces, to put back if the call stops
      const sheets = s.input.sheets.slice()
      applyLayout(s.input, checked, (name) => sheetIndex(s)(name)!)
      try {
        // merged areas added or removed: the arrays whose range they meet spill or stop; hidden rows: SUBTOTAL and AGGREGATE count again
        const keys = new Set<string>()
        for (let i = 0; i < s.input.sheets.length; i++) {
          const x = s.input.sheets[i]
          const now = new Set(x.merges ?? [])
          const moved = [...before[i]].filter((m) => !now.has(m)).concat([...now].filter((m) => !before[i].has(m)))
          if (moved.length) for (const k of yield* applyMerges(s, i, moved.map((m) => parseMerge(m)!))) keys.add(k)
          if (JSON.stringify(hiddenRows(x) ?? null) !== hidden[i]) for (const k of yield* applyHidden(s, i)) keys.add(k)
        }
        return changes(s, shown ? merged([...keys], shownChanges(s, shown)) : [...keys])
      } catch (e) {
        s.input.sheets.splice(0, sheets.length, ...sheets)
        stale = true
        throw e
      }
    },
    book() {
      return st().input
    },
    *explain(sheet: string, cell: string, o: ExplainOptions = {}) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const { row, col } = parseAddress(checkAddress(cell, 'cell'))
      const e = explainCell(s, i, row, col, o.depth ?? 3, o.maxNodes ?? 200, o.expand ?? false)
      return { ...e, text: explainText(e.root) }
    },
    *result() {
      const t0 = nowMs()
      return buildResult(yield* ready(), clock, t0, display)
    },
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
    setNames: (names) => now(() => core.setNames(names)),
    setClock: (clock) => now(() => core.setClock(clock)),
    restructure: (op) => now(() => core.restructure(op)),
    setStyles: (styles) => core.setStyles(styles),
    setLayout: (edits) => now(() => core.setLayout(edits)),
    book: () => core.book(),
    explain: (sheet, cell, o) => now(() => core.explain(sheet, cell, o)),
    result: () => now(() => core.result()),
    close: () => core.close(),
  }
}
