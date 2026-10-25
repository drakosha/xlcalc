/**
 * Live objects of a synchronous session's sheets and cells (`liveBook`, `docs/live-objects.md`): a cell reads the
 * session, an assignment is its edit, a batch is one `applyChanges`. One subscription to the session tells the
 * workbook's listeners: the book's all, a sheet's those of the changes that touched it, a cell's by a map of addresses,
 * so that an edit of cells costs the cells it changed, not the listeners; a structure operation wakes the cells of its
 * sheet whose address other cells may have come to.
 */
import type { CellTextOptions, CellTextResult, CellValue, LiveCell, LiveSheet, LiveBook, NumberFormat, Scalar, Session, SessionChange, SessionOperation, StylePatch } from './contract.js'
import { nameKey, parseCell, parseRange } from './a1.js'
import { InputError, what } from './input-error.js'
import { listenerOf, tell as tellListener, type Listener } from './session/subscribers.js'
import { strict } from './strict-args.js'

/** A watch: its own entry, so that one listener given twice is unsubscribed once. */
interface Watch {
  listener: Listener
}

/** The watches of a cell, with its row and column for a structure operation's area. */
interface CellWatches {
  row: number
  col: number
  watches: Set<Watch>
}

/** A sheet of the workbook while it lives: its name now, its object, its cells' objects and the watches. */
interface SheetState {
  name: string
  live: boolean
  sheet: LiveSheet
  cells: Map<string, WeakRef<LiveCell>>
  watches: Set<Watch>
  cellWatches: Map<string, CellWatches>
}

/** The workbook of each session: one subscription to it, the same objects for it. */
const books = new WeakMap<Session, LiveBook>()

/** A key of `cells` that may be the address of a cell, without `$` (`parseCell` checks it is on the sheet). */
const ADDRESS = /^[A-Za-z]{1,3}[1-9]\d{0,6}$/

/** A session of the synchronous kind: an asynchronous one answers `getBook()` with a promise. `getBook` changes nothing. */
function checkSession(session: unknown): Session {
  const s = session as Partial<Session> | null
  if (typeof s !== 'object' || s === null || typeof s.getBook !== 'function' || typeof s.subscribe !== 'function')
    throw new InputError('session', `expected a synchronous session (createSession), got ${what(session)}`)
  const book = s.getBook() as unknown as { then?: unknown } | null
  if (book && typeof book.then === 'function') {
    ;(book as Promise<unknown>).then(undefined, () => {})
    throw new InputError('session', 'expected a synchronous session (createSession), got an asynchronous one (createInlineCalc, createCalcPool)')
  }
  return s as Session
}

/** Adds a watch to `set`, answering the function that takes it away (nothing once `set` is cleared). */
function watchIn(set: Set<Watch>, listener: Listener, gone?: () => void): () => void {
  const w = { listener }
  set.add(w)
  return () => {
    if (set.delete(w) && !set.size) gone?.()
  }
}

/**
 * The live workbook of a synchronous session (`LiveBook`): the same object for the same session. An asynchronous
 * session throws `InputError`: its writes cannot return what they calculated.
 */
export function liveBook(session: Session): LiveBook {
  const s = checkSession(session)
  const known = books.get(s)
  if (known) return known

  const sheets = new Map<string, SheetState>()
  const bookWatches = new Set<Watch>()
  /** The assignments of a batch running, in order; null outside one. */
  let pending: SessionOperation[] | null = null
  const collected = new FinalizationRegistry<{ st: SheetState; address: string }>(({ st, address }) => {
    if (st.cells.get(address)?.deref() === undefined) st.cells.delete(address)
  })

  const alive = (st: SheetState): SheetState => {
    if (!st.live) throw new InputError('sheet', `the sheet "${st.name}" was deleted`)
    return st
  }

  /** An assignment: the session's call now, or an operation of the batch running. */
  function edit(op: SessionOperation): void {
    if (pending) pending.push(op)
    else if (op.op === 'setCells') s.setCells(op.edits)
    else if (op.op === 'patchCells') s.patchCells(op.patches)
    else if (op.op === 'clearCells') s.clearCells(op.ranges, op.what)
  }

  class Cell implements LiveCell {
    readonly #st: SheetState
    readonly address: string
    constructor(st: SheetState, address: string) {
      this.#st = st
      this.address = address
    }
    get sheet(): LiveSheet {
      return alive(this.#st).sheet
    }
    #view() {
      return s.getCell(alive(this.#st).name, this.address)
    }
    get value(): CellValue {
      return this.#view()?.value ?? null
    }
    set value(value: Scalar | { date: string }) {
      const sheet = alive(this.#st).name
      if (value === null) edit({ op: 'clearCells', ranges: [{ sheet, range: this.address }], what: 'contents' })
      else edit({ op: 'setCells', edits: [{ sheet, cell: this.address, input: { value } }] })
    }
    get text(): string | undefined {
      return this.#view()?.text
    }
    get formula(): string | undefined {
      return this.#view()?.formula
    }
    set formula(formula: string | null) {
      edit({ op: 'setCells', edits: [{ sheet: alive(this.#st).name, cell: this.address, input: { formula } }] })
    }
    get numberFormat(): NumberFormat | undefined {
      return this.#view()?.numberFormat
    }
    set numberFormat(numberFormat: NumberFormat | null) {
      edit({ op: 'patchCells', patches: [{ sheet: alive(this.#st).name, cell: this.address, set: { numberFormat } }] })
    }
    get style(): number | undefined {
      return this.#view()?.style
    }
    set style(style: number | StylePatch | null) {
      edit({ op: 'patchCells', patches: [{ sheet: alive(this.#st).name, cell: this.address, set: { style } }] })
    }
    type(text: string, opts?: CellTextOptions): CellTextResult {
      const sheet = alive(this.#st).name
      if (pending) throw new InputError('type', 'text typed into a cell cannot be part of a batch (applyChanges takes no typed text): type it before or after the batch')
      return s.setCellText(sheet, this.address, text, opts)
    }
    watch(listener: Listener): () => void {
      const st = alive(this.#st)
      const l = listenerOf('watch', listener)
      let w = st.cellWatches.get(this.address)
      if (!w) {
        const { row, col } = parseCell(this.address)!
        st.cellWatches.set(this.address, (w = { row, col, watches: new Set() }))
      }
      return watchIn(w.watches, l, () => {
        if (st.cellWatches.get(this.address) === w) st.cellWatches.delete(this.address)
      })
    }
  }
  for (const k of ['type', 'watch'] as const) Object.defineProperty(Cell.prototype, k, { value: strict(Cell.prototype[k], k) })

  /** The cell of `st` at `key`, a key of `cells`: the object held already, else a new one; undefined for no address. */
  function cellAt(st: SheetState, key: string): LiveCell | undefined {
    const held = st.cells.get(key)?.deref()
    if (held) return held
    if (!ADDRESS.test(key)) return undefined
    const address = key.toUpperCase()
    if (address !== key) return cellAt(st, address)
    const p = parseCell(address)
    if (!p) return undefined
    const c = new Cell(st, address)
    st.cells.set(address, new WeakRef(c))
    collected.register(c, { st, address })
    return c
  }

  function sheetOf(name: string): SheetState {
    const key = nameKey(name)
    const known = sheets.get(key)
    if (known) return known
    const found = s.getLayout().sheets.find((x) => nameKey(x.name) === key)
    if (!found) throw new InputError('name', `unknown sheet "${name}"`)
    const st: SheetState = { name: found.name, live: true, sheet: undefined as never, cells: new Map(), watches: new Set(), cellWatches: new Map() }
    const cells = new Proxy(Object.create(null) as Record<string, LiveCell>, {
      get: (_, key) => (typeof key === 'string' ? cellAt(alive(st), key) : undefined),
      has: (_, key) => typeof key === 'string' && cellAt(alive(st), key) !== undefined,
      set: () => false,
      defineProperty: () => false,
      deleteProperty: () => false,
    })
    st.sheet = Object.defineProperties(Object.create(null) as LiveSheet, {
      name: { get: () => alive(st).name, enumerable: true },
      cells: { get: () => (alive(st), cells) },
      cell: {
        value: strict(function cell(address: string) {
          const c = typeof address === 'string' ? cellAt(alive(st), address) : undefined
          if (!c) throw new InputError('address', `expected the address of a cell such as "B12", without $, got ${what(address)}`)
          return c
        }),
      },
      watch: { value: strict(function watch(listener: Listener) {
        return watchIn(alive(st).watches, listenerOf('watch', listener))
      }) },
    })
    sheets.set(key, st)
    return st
  }

  /** The watches of the cells of `st` in a rectangle: rows and columns from 1, bounds inclusive. */
  function wakeArea(st: SheetState, wake: Set<Watch>, top: number, left: number, bottom: number, right: number) {
    for (const w of st.cellWatches.values()) if (w.row >= top && w.row <= bottom && w.col >= left && w.col <= right) for (const x of w.watches) wake.add(x)
  }

  /** Follows a change: the sheets renamed and deleted, then the watches it touched, each listener called once. */
  function tell(change: SessionChange) {
    const wake = new Set<Watch>(bookWatches)
    const state = (name: string) => sheets.get(nameKey(name))
    const touch = (st: SheetState | undefined) => {
      if (st) for (const w of st.watches) wake.add(w)
      return st
    }
    for (const m of change.moved ?? []) {
      if (m.kind === 'renameSheet') {
        const st = state(m.sheet)
        if (!st) continue
        sheets.delete(nameKey(m.sheet))
        st.name = m.name
        sheets.set(nameKey(m.name), st)
        touch(st)
      } else if (m.kind === 'deleteSheet') {
        const st = state(m.sheet)
        if (!st) continue
        sheets.delete(nameKey(m.sheet))
        st.live = false
        for (const w of st.watches) wake.delete(w)
        for (const c of st.cellWatches.values()) for (const w of c.watches) wake.delete(w)
        st.watches.clear()
        st.cellWatches.clear()
      } else if (m.kind === 'moveSheet') touch(state(m.sheet))
      else if (m.kind === 'lines') {
        const st = touch(state(m.sheet))
        if (!st) continue
        const [top, left] = m.axis === 'rows' ? [m.at, 1] : [1, m.at]
        wakeArea(st, wake, top, left, Infinity, Infinity)
      } else if (m.kind === 'range') {
        const from = parseRange(m.from.range)
        const to = parseCell(m.to.cell)
        if (!from || !to) continue
        const src = touch(state(m.from.sheet))
        const dst = touch(state(m.to.sheet))
        if (src) wakeArea(src, wake, from.top, from.left, from.bottom, from.right)
        if (dst) wakeArea(dst, wake, to.row, to.col, to.row + from.bottom - from.top, to.col + from.right - from.left)
      }
    }
    if (change.styles) for (const st of sheets.values()) touch(st)
    for (const name of change.layout ?? []) touch(state(name))
    for (const r of change.repaint ?? []) touch(state(r.sheet))
    let name: string | undefined
    let st: SheetState | undefined
    for (const c of change.changes) {
      if (c.sheet !== name) {
        name = c.sheet
        st = touch(state(name))
      }
      const w = st?.cellWatches.get(c.cell)
      if (w) for (const x of w.watches) wake.add(x)
    }
    const called = new Set<Listener>()
    for (const { listener } of wake) {
      if (called.has(listener)) continue
      called.add(listener)
      tellListener(listener, change)
    }
  }
  s.subscribe(tell)

  const book: LiveBook = {
    session: s,
    sheet: strict(function sheet(name: string) {
      if (typeof name !== 'string') throw new InputError('name', `expected a sheet's name, got ${what(name)}`)
      return sheetOf(name).sheet
    }),
    batch: strict(function batch(fn: () => void) {
      if (typeof fn !== 'function') throw new InputError('batch', `fn must be a function, got ${what(fn)}`)
      const outer = pending === null
      const ops = (pending ??= [])
      const mark = ops.length
      try {
        const r = fn() as unknown as { then?: unknown } | undefined
        if (r && typeof r.then === 'function') throw new InputError('batch', 'fn must not return a promise: an assignment after an await is not part of the batch')
      } catch (e) {
        if (outer) pending = null
        else ops.length = mark
        throw e
      }
      if (!outer) return
      pending = null
      if (ops.length) s.applyChanges(ops)
    }),
    watch: strict(function watch(listener: Listener) {
      return watchIn(bookWatches, listenerOf('watch', listener))
    }),
  }
  books.set(s, book)
  return book
}
