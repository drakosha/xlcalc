/**
 * The engine: formulas are compiled (`compile.ts`) to closures or, in the `eval` mode, to JS text with one
 * `new Function` per batch; cells are computed over the reference graph.
 *
 * Each formula is its own function (editing a formula in a session compiles only that formula). A full
 * computation follows the topological order; editing a value recomputes only the dependent cells.
 *
 * No platform imports, so the engine runs outside Node as well.
 */
import type { BookInput, CachedValues, CellInput, CellReplace, CellValue, CheckCode, ErrorCode, Execution, Externals, NameIssue, SheetInput } from './contract.js'
import { makeCompiler, PRELUDE, type CodegenContext, type Compiler, type FormulaCode } from './compile.js'
import { changed, concerned, evaluateAll, put, recompute, tracker, write, type Tracker } from './evaluate.js'
import { count, due, run, step, type Steps } from './exec/budget.js'
import { FUNCTIONS } from './functions.js'
import { readersOf, rebuildGraph } from './graph.js'
import { epochOf, ymd } from './dates.js'
import { builtAnew, heldValue, putHeld } from './held.js'
import { hiddenRows } from './layout/hidden.js'
import { localeBooleans } from './locale/registry.js'
import { rulesOf } from './locale/rules.js'
import { markingClock } from './locale/rules/data.js'
import { addressOrder, bookScalar, formatAddress, MAX_COLS, parseAddress, placeOf } from './model.js'
import { makeRuntime, type Runtime, type RuntimeContext, type SheetData } from './runtime.js'
import { heapCaches, RangeCache, type CacheBudget } from './range-cache.js'
import { RangeAggregates } from './range-aggregate.js'
import { AreaIndex } from './area-index.js'
import { expandLinkNames, linkedSheet, linkSheets, strictCells, strictLinks, type LinkState } from './links.js'
import { linksIn, markedLinks } from './syntax/links.js'
import { nameFor, prepareNames, type PreparedName } from './names.js'
import { endSpill, spillRef } from './spill/place.js'
import { newSpills, type Area, type Spills } from './spill/state.js'
import { meets, type Rect } from './structure/area.js'
import { arrayArea, holds } from './array-areas.js'
import { lookAlone, looksTaken, putInput, settle } from './looks/sheet.js'
import { quoteSheet, sheetKey } from './syntax.js'
import { prepare, prepareShaped, type Prepared, type Shapes } from './prepare.js'
import { jsBackend } from './exec/js.js'
import { walk } from './exec/walk.js'
import { specRect, tableKey, tableNames, tablesOf, type Tables } from './tables/model.js'
import { E, isErr, type V } from './values.js'

export interface Formula extends FormulaCode {
  id: number
  s: number
  r: number
  c: number
  /** Canonical text. */
  text: string
  /** An array formula of the old kind: its area (`CellInput.array`, `spill/fixed.ts`). */
  area?: Rect
}

export interface Invalid {
  code: CheckCode
  message: string
  at: number
  error: ErrorCode
}

export interface EngineOptions {
  /** The clock moment, a serial number of the 1900 date system (UTC). */
  now: number
  seed: number
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /** What the range caches may hold, shared with other books: those of this heap by default. */
  cache?: CacheBudget
  /** Whether the state may take a sheet's cells as they are, no one else reading or changing them after: else it takes a copy. */
  own?: (cells: Record<string, CellInput>) => boolean
  /** The addresses of cells objects in their order, listed already (`checkBook`): the state reads those it takes by them. */
  keys?: Map<object, string[]>
  /** The values the book's file stores (`ComputeOptions.cached`), for the formulas a circular reference stops that hold none. */
  cached?: CachedValues
  /** What a formula showed before a session's book was built again, for one a circular reference stops now that holds none (`held.ts`). */
  prior?: (f: Formula) => V | undefined
  /** Gets the sheets whose cells the state copied to change a held value: the book given holds other inputs there. */
  rewritten?: Set<number>
}

export interface EngineState {
  /** Temporary value journal of an impact preview. */
  capture?: Tracker
  input: BookInput
  sheetNames: string[]
  sheetIdx: Map<string, number>
  names: PreparedName[]
  nameIssues: NameIssue[]
  tables: Tables
  externalFunctions: string[]
  ctx: RuntimeContext
  rt: Runtime
  compiler: Compiler
  /** id → formula; a hole after deletion. */
  formulas: (Formula | undefined)[]
  /** Reusable formula slots; helpers are rebuilt after every structural edit. */
  free: number[]
  /** `s:r:c` → formula id. */
  at: Map<number, number>
  invalid: Map<string, Invalid>
  /** Dependency vertices: formula ids first, shared range helpers after the formula slots. */
  succ: number[][]
  /** The formulas reading a spilled array's range, by the array's formula (successors besides `succ`). */
  spillSucc: Map<number, number[]>
  /** Position in the topological order; -1 for cells in a cycle and cells depending on them. */
  topo: number[]
  order: number[]
  /** A walk of the graph from changed formulas (`ordered`): a vertex it reached holds its mark, so no set is filled. */
  marks: Uint32Array
  mark: number
  /** How many formulas are arrays of the old kind (`Formula.area`): edits of a book with none look at no formula for them. */
  arrays: number
  /** Formula references indexed by sheet and column, with a separate bucket for wide ranges. */
  listeners: AreaIndex<Area & { id: number }>
  /** Formulas that depend on volatile functions or on `clocked` formulas (directly or through others). */
  volatile: Set<number>
  /** Formulas whose latest computation read the clock's year (date text without a year, by the book's locale). */
  clocked: Set<number>
  /** The formulas in cycles the last iterative calculation computed (`BookInput.iteration`). */
  iterated: number[]
  /** The values of the formulas a circular reference stops that hold one (`CellInput.held`, `held.ts`), by formula. */
  held: Map<number, V>
  /**
   * What a formula a circular reference stops takes when the cycle comes into its way and it holds no value (`held.ts`):
   * what a function gives while the state is built or a book is put back (0 where it gives none); 0 for the formulas the
   * edits being applied entered, what the others show; what it shows when none is set.
   */
  anew?: Set<number> | ((f: Formula) => V | undefined)
  /** Formulas computed this time that hold a value: out of a cycle's way once the computation settles, they drop it. */
  outside: Set<number>
  /** Where a caller of `applyEdits` gets the inputs the engine changed besides the edits (a held value taken or dropped), each cell's first, by `s!A1`. */
  journal?: Map<string, CellReplace>
  /** While the state is built: the sheets whose cells it copied to change a held value (`EngineOptions.rewritten`). */
  cow?: Set<number>
  spills: Spills
  /** Input cells that hold no value or formula (a style, a format, a note alone), by sheet (row · MAX_COLS + column): a sheet's extent counts them. */
  bare: Map<number, Set<number>>
  /** The linked workbooks' sheets in the runtime after the book's own, and their names (`links.ts`). */
  links: LinkState
}

export const key = (s: number, r: number, c: number) => `${s}:${r}:${c}`

/** Marks or unmarks a cell whose formula calls SUBTOTAL or AGGREGATE. */
function mark(ctx: RuntimeContext, s: number, r: number, c: number, on: boolean) {
  let marks = ctx.nested.get(s)
  if (!marks) ctx.nested.set(s, (marks = new Set()))
  if (on) marks.add(r * MAX_COLS + c)
  else marks.delete(r * MAX_COLS + c)
}

/** The hidden rows of sheet `s` that SUBTOTAL and AGGREGATE read, from its input. */
function setHidden(ctx: RuntimeContext, s: number, sheet: SheetInput) {
  const h = hiddenRows(sheet)
  if (h) ctx.hidden.set(s, h)
  else ctx.hidden.delete(s)
}

/** Marks whether a cell's input is one without a value or formula (`EngineState.bare`). */
function markBare(state: EngineState, s: number, r: number, c: number, input: CellInput | null) {
  let bare = state.bare.get(s)
  if (input && !holds(input)) {
    if (!bare) state.bare.set(s, (bare = new Set()))
    bare.add(r * MAX_COLS + c)
  } else bare?.delete(r * MAX_COLS + c)
}

/**
 * Sets a cell's input (null removes it), and whether it is one without a value or formula; an input of a look alone goes
 * to the sheet's looks (`putInput`), which take it when the edits are settled (`settled`).
 */
export function setInput(state: EngineState, s: number, addr: string, r: number, c: number, input: CellInput | null) {
  putInput(state.input.sheets[s], addr, r, c, input)
  markBare(state, s, r, c, input && !lookAlone(input) ? input : null)
}

/** The index of the sheet named `name` in any case; a name spelled as the book spells it is found without its key, which upper-cases it. */
export function sheetIndexOf(state: EngineState, name: string): number | undefined {
  const i = state.sheetNames.indexOf(name)
  return i >= 0 ? i : state.sheetIdx.get(sheetKey(name))
}

/** The looks of the sheets edited cells were written to, settled (`settle`): the book is as the engine takes it again. */
export function settled(state: EngineState, edits: { sheet: string }[]) {
  for (const i of new Set(edits.map((e) => sheetIndexOf(state, e.sheet)))) if (i !== undefined) settle(state.input.sheets[i])
}

function emptySheet(name: string): SheetData {
  return { name, cols: [], rows: 0, fills: [] }
}


// ---------- preparation ----------

export const codegenContext = (state: EngineState, s: number, area?: Rect): CodegenContext => ({
  sheet: (name) => (name === null ? s : state.sheetIdx.get(sheetKey(name))!),
  name: (name, scope) => nameFor(state.names, scope, name),
  table: (n, r) => {
    const t = n.table === null ? undefined : state.tables.byName.get(tableKey(n.table))
    const x = t ? specRect(t, n.spec, r) : 'REF'
    return typeof x === 'string' ? x : { ...x, s: t!.s }
  },
  link: (book, sheet) => linkedSheet(state.links, book, sheet),
  ...(area ? { area } : {}),
})

/** The area of a formula cell's array formula of the old kind, if it is one. */
const areaOf = (cell: CellInput, row: number, col: number) => (cell.array === undefined ? undefined : (arrayArea(formatAddress(row, col), cell.array) ?? undefined))

/**
 * An array formula of the old kind that does not prepare, as its error over its area: Excel shows the error in every
 * cell of the area, keeps the area whole and spills nothing into it (the spill set `error-areas-held`).
 */
const erring = (cell: CellInput, invalid: Invalid): Extract<Prepared, { ok: true }> => ({ ok: true, text: cell.formula!, ast: { k: 'err', code: invalid.error } })

/** A formula cell compiled: its code, and the area of an old-kind array formula. */
function compiled(state: EngineState, id: number, s: number, row: number, col: number, cell: CellInput, p: Extract<Prepared, { ok: true }>): Formula {
  const area = areaOf(cell, row, col)
  const g = state.compiler.formula(codegenContext(state, s, area), p.ast, s, row, col)
  return { id, s, r: row, c: col, text: p.text, ...g, ...(area ? { area } : {}) }
}

/** Graph-relevant metadata of a scalar replacement; names, arrays and spans retain the rebuild path. */
function sameDependencies(a: Formula, b: Formula): boolean {
  return !a.area && !b.area && !a.spans && !b.spans && a.volatile === b.volatile && a.nested === b.nested && a.external === b.external
    && a.refs.length === b.refs.length && a.refs.every((g, i) => {
      const h = b.refs[i]
      return g.s === h.s && g.r1 === h.r1 && g.c1 === h.c1 && g.r2 === h.r2 && g.c2 === h.c2
    })
}

// ---------- assembly ----------

/** A book compiled and computed at once. */
export const createState = (book: BookInput, opts: EngineOptions): EngineState => run(buildState(book, opts))

/** A book compiled and computed, in steps: between two formulas compiled or computed. */
export function* buildState(book: BookInput, opts: EngineOptions): Steps<EngineState> {
  const sheetNames = book.sheets.map((s) => s.name)
  const sheetIdx = new Map(sheetNames.map((n, i) => [sheetKey(n), i]))
  if (sheetIdx.size !== sheetNames.length) throw new Error('duplicate sheet name')
  // the state's own cells (edits change them), copied unless they are its own already, the entries of a look alone in the
  // sheet's looks (a structure operation over looks gives them back so); the layout and the styles are shared
  // no closure here takes \`opts\`: the state's closures share this scope, and \`opts.prior\` holds the state built before
  const { own, keys } = opts
  const input: BookInput = { ...book, sheets: book.sheets.map((x) => looksTaken(x)).map((s) => ({ ...s, cells: own?.(s.cells) ? s.cells : { ...s.cells } })) }
  const externalFunctions = (opts.externals?.functions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f])
  const tables = tablesOf(input)
  const prepared = prepareNames(sheetIdx, input.names ?? [], externalFunctions, tableNames(tables), localeBooleans(book.locale), book.externalBooks)
  const linked = linkSheets(book.externalBooks, sheetNames.length)
  const names = prepared.names.map((n) => ({ ...n, ast: expandLinkNames(n.ast, linked.links, book.externalBooks) }))
  const epoch = epochOf(book.dateSystem)
  // the clock's year completes date text without a year (ru-RU `1.5`); reading it marks the formula being computed
  const clock = markingClock(() => ymd(Math.floor(ctx.now) + epoch).y)
  const cache = opts.cache ?? heapCaches
  const ctx: RuntimeContext = {
    sheets: [...sheetNames.map(emptySheet), ...linked.data],
    now: opts.now - epoch,
    seed: opts.seed,
    rules: rulesOf(book.locale, clock, epoch),
    externals: opts.externals?.values ?? {},
    cache,
    flat: new RangeCache(cache),
    index: new RangeCache(cache),
    lower: new RangeCache(cache),
    selections: new RangeCache(cache),
    aggregates: new RangeAggregates(cache),
    nested: new Map(),
    hidden: new Map(),
    spillRef: (s, r, c) => spillRef(state, s, r, c),
  }
  input.sheets.forEach((sh, s) => setHidden(ctx, s, sh))
  const rt = makeRuntime(ctx)
  const state: EngineState = {
    input,
    sheetNames,
    sheetIdx,
    names,
    nameIssues: prepared.issues,
    tables,
    externalFunctions,
    ctx,
    rt,
    compiler: makeCompiler(rt, opts.execution ?? 'closures'),
    formulas: [],
    free: [],
    at: new Map(),
    invalid: new Map(),
    succ: [],
    spillSucc: new Map(),
    topo: [],
    order: [],
    marks: new Uint32Array(0),
    mark: 0,
    arrays: 0,
    listeners: new AreaIndex(),
    volatile: new Set(),
    clocked: new Set(),
    iterated: [],
    held: new Map(),
    outside: new Set(),
    cow: opts.rewritten ?? new Set(),
    spills: newSpills(),
    bare: new Map(),
    links: linked.links,
  }
  const pending: Formula[] = []
  // the formulas of one shape are prepared once while this state is built
  const shapes: Shapes = new Map()
  for (let s = 0; s < input.sheets.length; s++) {
    const sheet = input.sheets[s]
    // the cells in the order of their entries, listed already where the caller has them
    for (const addr of keys?.get(sheet.cells) ?? Object.keys(sheet.cells)) {
      const cellIn = sheet.cells[addr]
      const place = addressOrder(addr), col = place % MAX_COLS, row = (place - col) / MAX_COLS
      if (holds(cellIn)) state.spills.occupied.set(s, row, col, true, true)
      // a held value means nothing but in a formula's cell
      if (cellIn.formula === undefined) {
        if (cellIn.held !== undefined) putHeld(state, s, addr, undefined)
        const v = bookScalar(cellIn.value, epoch)
        if (v !== null) put(ctx.sheets[s], row, col, v)
        else markBare(state, s, row, col, cellIn)
        continue
      }
      if (due()) yield
      step()
      count(pending.length + state.invalid.size + 1)
      let p = prepareShaped(state, shapes, s, row, col, cellIn.formula)
      if (!p.ok) {
        state.invalid.set(key(s, row, col), p.invalid)
        if (cellIn.array !== undefined) p = erring(cellIn, p.invalid)
        else {
          put(ctx.sheets[s], row, col, E[p.invalid.error])
          if (cellIn.held !== undefined) putHeld(state, s, addr, undefined)
          continue
        }
      }
      const f = compiled(state, pending.length, s, row, col, cellIn, p)
      if (cellIn.held !== undefined) state.held.set(f.id, heldValue(cellIn.held))
      pending.push(f)
      if (f.area) state.arrays++
      if (f.nested) mark(ctx, s, row, col, true)
    }
  }
  state.compiler.link(pending)
  state.formulas = pending
  if (state.links.strict.some(Boolean)) for (const [s, r, c] of strictCells(state.links, ctx.sheets, pending.flatMap((f) => f.refs))) put(ctx.sheets[s], r, c, E.REF)
  for (const f of state.formulas) state.at.set(placeOf(f!.s, f!.r, f!.c), f!.id)
  rebuildGraph(state)
  state.anew = builtAnew(state, opts.cached, opts.prior)
  yield* evaluateAll(state)
  state.anew = state.cow = undefined
  return state
}

/**
 * Edits applied atomically, one recomputation of dependents; returns the changed cells' places (`placeOf`).
 * `withVolatile`: the clock has changed (live session), volatile cells are recomputed in the same pass. `restore`: the
 * edits put inputs back (an undo, a preview taken back), so a formula a circular reference stops holds what its input
 * gives, as a full computation of the book gives (`held.ts`).
 */
export function* applyEdits(state: EngineState, edits: CellReplace[], withVolatile = false, restore = false): Steps<number[]> {
  const track = tracker()
  const dirty = new Set<number>()
  const fresh: Formula[] = []
  // the edited cells and the ranges of arrays whose formula is gone
  const touched: Area[] = []
  let structural = false
  let graphChanged = false
  for (const e of edits) {
    const s = sheetIndexOf(state, e.sheet)
    if (s === undefined) throw new Error(`unknown sheet ${e.sheet}`)
    const { row, col } = parseAddress(e.cell)
    const k = key(s, row, col), n = placeOf(s, row, col)
    setInput(state, s, e.cell, row, col, e.input)
    state.spills.occupied.set(s, row, col, holds(e.input ?? undefined))
    const old = state.at.get(n)
    const previous = old === undefined ? undefined : state.formulas[old]
    const clocked = old !== undefined && state.clocked.has(old)
    const spilled = old !== undefined && (state.spills.claims.has(old) || state.spills.on.has(old) || state.spills.wait.has(old))
    if (old !== undefined) {
      if (state.formulas[old]!.nested) mark(state.ctx, s, row, col, false)
      if (state.formulas[old]!.area) state.arrays--
      endSpill(state, state.formulas[old]!, track, false)
      state.held.delete(old)
      state.formulas[old] = undefined
      state.free.push(old)
      state.clocked.delete(old)
      state.at.delete(n)
      structural = true
    }
    state.invalid.delete(k)
    // An atomic column batch touches one contiguous area, without querying all its readers per cell.
    const last = touched.at(-1)
    if (last && last.s === s && last.c1 === col && last.c2 === col && row >= last.r1 - 1 && row <= last.r2 + 1) {
      last.r1 = Math.min(last.r1, row)
      last.r2 = Math.max(last.r2, row)
    } else touched.push({ s, r1: row, c1: col, r2: row, c2: col })
    const input: CellInput | null = e.input
    if (input?.formula !== undefined) {
      let p = prepare(state, s, row, col, input.formula)
      if (!p.ok) {
        state.invalid.set(k, p.invalid)
        if (input.array !== undefined) p = erring(input, p.invalid)
        else {
          graphChanged ||= !!previous
          write(state, s, row, col, E[p.invalid.error], track)
          if (input.held !== undefined) putHeld(state, s, e.cell, undefined)
          continue
        }
      }
      const f = compiled(state, state.free.pop() ?? state.formulas.length, s, row, col, input, p)
      if (input.held !== undefined) state.held.set(f.id, heldValue(input.held))
      if (f.nested) mark(state.ctx, s, row, col, true)
      if (f.area) state.arrays++
      state.formulas[f.id] = f
      state.at.set(n, f.id)
      fresh.push(f)
      graphChanged ||= !previous || previous.id !== f.id || clocked || spilled || state.spills.claims.size > 0 || state.spills.wait.size > 0 || !sameDependencies(previous, f)
      structural = true
    } else {
      graphChanged ||= !!previous
      if (input?.held !== undefined) putHeld(state, s, e.cell, undefined)
      write(state, s, row, col, input ? bookScalar(input.value, state.ctx.rules.epoch) : null, track)
    }
  }
  settled(state, edits)
  count(state.at.size + state.invalid.size)
  if (structural) {
    const size = state.formulas.length
    while (state.formulas.length && !state.formulas.at(-1)) state.formulas.pop()
    if (state.formulas.length < size) state.free = state.free.filter((id) => id < state.formulas.length)
  }
  state.compiler.link(fresh)
  // the formulas a cycle stopped are computed again, in case it is broken
  if (structural && !state.input.iteration) for (const f of state.formulas) if (f && state.topo[f.id] < 0) dirty.add(f.id)
  if (graphChanged) rebuildGraph(state)
  if (fresh.length && state.links.books.length) touched.push(...enteredLinks(state, fresh))
  // formulas reading the cells, arrays spilled over them or waiting for them
  for (const x of [...touched, ...state.spills.touched.splice(0)]) concerned(state, x, dirty)
  for (const f of fresh) dirty.add(f.id)
  if (withVolatile) for (const id of state.volatile) dirty.add(id)
  state.anew = restore ? () => undefined : fresh.length ? new Set(fresh.map((f) => f.id)) : undefined
  try {
    yield* recompute(state, dirty, track, !structural)
  } finally {
    state.anew = undefined
  }
  return changed(state, track)
}

/**
 * Formulas entered that refer to linked books: Excel cannot read those books and marks every sheet of them
 * (`markedLinks`), so the cells their stored values do not hold become `#REF!` where formulas read them. They are
 * written without a change for the host (a linked sheet is none of the book's); the areas written, whose readers compute again.
 */
function enteredLinks(state: EngineState, fresh: Formula[]): Area[] {
  const ks = new Set(fresh.flatMap((f) => linksIn(f.text)))
  if (!ks.size) return []
  const books = markedLinks(state.input.externalBooks, ks)
  const marked = books !== state.input.externalBooks
  if (marked) {
    state.input.externalBooks = books
    strictLinks(state.links, books!, ks)
  }
  const refs = (marked ? state.formulas.filter((f): f is Formula => !!f) : fresh).flatMap((f) => f.refs)
  const out: Area[] = []
  for (const [s, r, c] of strictCells(state.links, state.ctx.sheets, refs)) if (write(state, s, r, c, E.REF)) out.push({ s, r1: r, c1: c, r2: r, c2: c })
  return out
}

/** The clock's moment (a serial number of the 1900 date system) in the book's date system. */
export function setNow(state: EngineState, now: number, seed: number) {
  state.ctx.now = now - state.ctx.rules.epoch
  state.ctx.seed = seed
}

/** A new clock: recomputes volatile cells and their dependents. */
export function* applyClock(state: EngineState, now: number, seed: number): Steps<number[]> {
  setNow(state, now, seed)
  const track = tracker()
  yield* recompute(state, new Set(state.volatile), track)
  return changed(state, track)
}

/** Merged areas of a sheet were added or removed (`rects`): the arrays whose range meets them are computed again. */
export function* applyMerges(state: EngineState, s: number, rects: Rect[]): Steps<number[]> {
  const track = tracker()
  const dirty = new Set<number>()
  state.spills.merges.delete(s)
  for (const m of [state.spills.on, state.spills.wait]) for (const [id, x] of m) if (x.s === s && rects.some((r) => meets(r, x))) dirty.add(id)
  yield* recompute(state, dirty, track)
  return changed(state, track)
}

/**
 * Hidden rows or the filter mode of sheet `s` changed: the formulas calling SUBTOTAL or AGGREGATE over the rows hidden
 * before or now are computed again.
 */
export function* applyHidden(state: EngineState, s: number): Steps<number[]> {
  const before = state.ctx.hidden.get(s)?.spans ?? []
  setHidden(state.ctx, s, state.input.sheets[s])
  const track = tracker()
  const readers = new Set<number>()
  for (const [r1, r2] of [...before, ...(state.ctx.hidden.get(s)?.spans ?? [])]) readersOf(state, { s, r1, r2, c1: 0, c2: MAX_COLS - 1 }, readers)
  yield* recompute(state, new Set([...readers].filter((id) => state.formulas[id]!.nested)), track)
  return changed(state, track)
}

/** The value of a cell (null when empty), a spilled one too. */
export function valueAt(state: EngineState, s: number, r: number, c: number): V {
  return state.rt.cell(s, r, c)
}

export function toCellValue(v: V): CellValue {
  return isErr(v) ? { error: v.code } : v
}

// ---------- debugging and explain ----------

/** Generated code in evaluation order, for reading (not part of the contract). */
export function debugSource(state: EngineState): string {
  const lines = [PRELUDE]
  const ids = [...state.order, ...state.formulas.filter((f) => f && state.topo[f.id] < 0).map((f) => f!.id)]
  for (const id of ids) {
    const f = state.formulas[id]!
    let source = f.source
    if (source === undefined) {
      const p = prepare(state, f.s, f.r, f.c, f.text)
      if (p.ok) {
        const b = jsBackend()
        source = b.source(walk(codegenContext(state, f.s, f.area), p.ast, f.s, f.r, f.c, b).code)
      }
    }
    lines.push(`// ${quoteSheet(state.sheetNames[f.s])}!${formatAddress(f.r, f.c)} = ${f.text}${state.topo[id] < 0 ? '  (cycle)' : ''}`, source ?? '')
  }
  return lines.join('\n')
}
