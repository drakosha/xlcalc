/**
 * A session over a book: edits are applied atomically with one recomputation of dependents. A live
 * session (`live`) takes a fresh clock and a new seed on every recomputation; its current clock
 * (`clock`) is what a saved result records, and a full computation with it gives the same hash.
 */
import type { BookInput, CellChange, CellEdit, CellStyle, Clock, Explanation, ExplainOptions, LayoutEdit, NameInput, Session, SessionOptions, SessionRestructureResult, StructureOp } from './contract.js'
import { applyClock, applyEdits, applyMerges, createState, toCellValue, valueAt, type EngineState } from './engine.js'
import { explainCell, explainText } from './explain.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkClock, checkEdits, checkLayoutEdits, checkNames, checkSessionOptions, checkSheet, checkStyles } from './input.js'
import { applyLayout, checkStylesUsed, contentEdits } from './session/book.js'
import { movedChanges } from './session/moved.js'
import { parseMerge } from './layout/merges.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, clockSerial, outValue, spillFields } from './result.js'
import { applyOp, checkedMapCell } from './structure/index.js'
import { sheetKey } from './syntax.js'

/**
 * Opens a session. Without options it keeps the clock of the moment it was opened (with a random seed)
 * until `setClock`; `{ live: true }` takes a fresh clock on every recomputation.
 */
export function openSession(book: BookInput, opts?: SessionOptions): Session {
  const o = checkSessionOptions(opts)
  checkBook(book)
  const live = o.live === true
  let clock: Clock = 'clock' in o && o.clock ? { ...o.clock } : currentClock()
  const { externals, execution } = o
  const build = (b: BookInput) => createState(b, { now: clockSerial(clock), seed: clock.seed, externals, execution })
  let state: EngineState | null = build(book)
  const st = () => {
    if (!state) throw new Error('session is closed')
    return state
  }
  const sheetIndex = (s: EngineState) => (name: string) => s.sheetIdx.get(sheetKey(name))
  const styleCount = (s: EngineState) => s.input.styles?.length ?? 0
  const changes = (s: EngineState, keys: string[]): CellChange[] =>
    keys.map((k) => {
      const [i, r, c] = k.split(':').map(Number)
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value: outValue(toCellValue(valueAt(s, i, r, c))), ...spillFields(s, i, r, c) }
    })
  const tick = (s: EngineState) => {
    clock = currentClock()
    s.ctx.now = clockSerial(clock)
    s.ctx.seed = clock.seed
  }
  return {
    get clock() {
      return { ...clock }
    },
    setCells(edits: CellEdit[]) {
      const s = st()
      checkEdits(edits, sheetIndex(s), styleCount(s))
      // edits that keep a cell's value and formula change its input alone
      const loud = contentEdits(s, edits)
      if (!loud.length) return []
      if (live) tick(s)
      return changes(s, applyEdits(s, loud, live))
    },
    setNames(names: NameInput[]) {
      const old = st()
      checkNames(names, 'names')
      if (live) tick(old)
      // names change the whole graph: the book is rebuilt, changes are collected over all cells (none moved)
      const next = build({ ...old.input, names })
      const keys = movedChanges(old, next, (ref) => ref)
      state = next
      return changes(next, keys)
    },
    setClock(next: Clock) {
      if (live) throw new Error('a live session takes its clock from the system')
      const s = st()
      clock = { ...checkClock(next, 'clock') }
      return changes(s, applyClock(s, clockSerial(clock), clock.seed))
    },
    restructure(op: StructureOp): SessionRestructureResult {
      const old = st()
      const r = applyOp(old.input, op)
      if (!r.ok) return r
      if (live) tick(old)
      // the simple way: the state is built again from the new book
      const next = build(r.book)
      const keys = movedChanges(old, next, r.mapCell)
      state = next
      return { ok: true, book: next.input, mapCell: checkedMapCell(old.input, r.mapCell), changes: changes(next, keys) }
    },
    setStyles(styles: CellStyle[]) {
      const s = st()
      checkStylesUsed(s.input, checkStyles(styles, 'styles').length)
      s.input.styles = styles
    },
    setLayout(edits: LayoutEdit[]) {
      const s = st()
      const before = s.input.sheets.map((x) => new Set(x.merges ?? []))
      applyLayout(s.input, checkLayoutEdits(edits, sheetIndex(s), styleCount(s)), (name) => sheetIndex(s)(name)!)
      // merged areas added or removed: the arrays whose range they meet spill or stop
      const keys = new Set<string>()
      s.input.sheets.forEach((x, i) => {
        const now = new Set(x.merges ?? [])
        const moved = [...before[i]].filter((m) => !now.has(m)).concat([...now].filter((m) => !before[i].has(m)))
        if (moved.length) for (const k of applyMerges(s, i, moved.map((m) => parseMerge(m)!))) keys.add(k)
      })
      return changes(s, [...keys])
    },
    book() {
      return st().input
    },
    explain(sheet: string, cell: string, o: ExplainOptions = {}): Explanation {
      const s = st()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const { row, col } = parseAddress(checkAddress(cell, 'cell'))
      const e = explainCell(s, i, row, col, o.depth ?? 3, o.maxNodes ?? 200, o.expand ?? false)
      return { ...e, text: explainText(e.root) }
    },
    result() {
      const t0 = nowMs()
      return buildResult(st(), clock, t0)
    },
    close() {
      state = null
    },
  }
}
