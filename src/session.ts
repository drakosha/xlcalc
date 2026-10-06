/**
 * A session over a book: edits are applied atomically with one recomputation of dependents. A live
 * session (`live`) takes a fresh clock and a new seed on every recomputation; its current clock
 * (`clock`) is what a saved result records, and a full computation with it gives the same hash.
 */
import type { BookInput, CellChange, CellEdit, CellStyle, Clock, Explanation, ExplainOptions, LayoutEdit, NameInput, Session, SessionOptions, SessionRestructureResult, StructureOp } from './contract.js'
import { applyClock, applyEdits, createState, explainCell, explainText, toCellValue, valueAt, type EngineState } from './engine.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkClock, checkEdits, checkLayoutEdits, checkNames, checkSessionOptions, checkSheet, checkStyles } from './input.js'
import { applyLayout, checkStylesUsed, contentEdits } from './session/book.js'
import { movedChanges } from './session/moved.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, clockSerial, outValue } from './result.js'
import { applyOp } from './structure/index.js'
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
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value: outValue(toCellValue(valueAt(s, i, r, c))) }
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
      const next = build({ ...old.input, names })
      // names change the whole graph: the book is rebuilt, changes are collected over all cells
      const keys: string[] = []
      next.input.sheets.forEach((_, i) => {
        const cols = Math.max(old.ctx.sheets[i].cols.length, next.ctx.sheets[i].cols.length)
        const rows = Math.max(old.ctx.sheets[i].rows, next.ctx.sheets[i].rows)
        for (let c = 0; c < cols; c++)
          for (let r = 0; r < rows; r++) {
            const a = toCellValue(valueAt(old, i, r, c))
            const b = toCellValue(valueAt(next, i, r, c))
            if (JSON.stringify(a) !== JSON.stringify(b)) keys.push(`${i}:${r}:${c}`)
          }
      })
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
      return { ok: true, book: next.input, mapCell: r.mapCell, changes: changes(next, keys) }
    },
    setStyles(styles: CellStyle[]) {
      const s = st()
      checkStylesUsed(s.input, checkStyles(styles, 'styles').length)
      s.input.styles = styles
    },
    setLayout(edits: LayoutEdit[]) {
      const s = st()
      applyLayout(s.input, checkLayoutEdits(edits, sheetIndex(s), styleCount(s)), (name) => s.sheetIdx.get(sheetKey(name))!)
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
