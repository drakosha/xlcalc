/**
 * The engine: formulas are compiled (`compile.ts`) to closures or, in the `eval` mode, to JS text with one
 * `new Function` per batch; cells are computed over the reference graph.
 *
 * Each formula is its own function (editing a formula in a session compiles only that formula). A full
 * computation follows the topological order; editing a value recomputes only the dependent cells.
 *
 * No platform imports, so the engine runs outside Node as well.
 */
import type { BookInput, CellInput, CellReplace, CellValue, CheckCode, ErrorCode, Execution, Externals, NameIssue, SheetInput } from './contract.js'
import { makeCompiler, PRELUDE, type CodegenContext, type Compiler, type FormulaCode } from './compile.js'
import { changed, concerned, evaluateAll, put, recompute, tracker, write } from './evaluate.js'
import { count, due, run, step, type Steps } from './exec/budget.js'
import { FUNCTIONS } from './functions.js'
import { readersOf, rebuildGraph } from './graph.js'
import { epochOf, ymd } from './dates.js'
import { hiddenRows } from './layout/hidden.js'
import { localeBooleans } from './locale/registry.js'
import { rulesOf } from './locale/rules.js'
import { markingClock } from './locale/rules/data.js'
import { bookScalar, formatAddress, MAX_COLS, parseAddress } from './model.js'
import { makeRuntime, type Runtime, type RuntimeContext, type SheetData } from './runtime.js'
import { heapCaches, RangeCache, type CacheBudget } from './range-cache.js'
import { RangeAggregates } from './range-aggregate.js'
import { AreaIndex } from './area-index.js'
import { nameFor, prepareNames, type PreparedName } from './names.js'
import { endSpill, spillRef } from './spill/place.js'
import { newSpills, type Area, type Spills } from './spill/state.js'
import { meets, type Rect } from './structure/area.js'
import { arrayArea, holds } from './array-areas.js'
import { printFormula, quoteSheet, sheetKey, validate, type Ast } from './syntax.js'
import { specRect, tableAt, tableKey, tableNames, tablesOf, type Tables } from './tables/model.js'
import { CErr, E, isErr, type V } from './values.js'

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

interface Invalid {
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
}

export interface EngineState {
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
  /** `s:r:c` → formula id. */
  at: Map<string, number>
  invalid: Map<string, Invalid>
  /** Dependency vertices: formula ids first, shared range helpers after the formula slots. */
  succ: number[][]
  /** The formulas reading a spilled array's range, by the array's formula (successors besides `succ`). */
  spillSucc: Map<number, number[]>
  /** Position in the topological order; -1 for cells in a cycle and cells depending on them. */
  topo: number[]
  order: number[]
  /** Formula references indexed by sheet and column, with a separate bucket for wide ranges. */
  listeners: AreaIndex<Area & { id: number }>
  /** Formulas that depend on volatile functions or on `clocked` formulas (directly or through others). */
  volatile: Set<number>
  /** Formulas whose latest computation read the clock's year (date text without a year, by the book's locale). */
  clocked: Set<number>
  /** The formulas in cycles the last iterative calculation computed (`BookInput.iteration`). */
  iterated: number[]
  spills: Spills
  /** Input cells that hold no value or formula (a style, a format, a note alone), by sheet (row · MAX_COLS + column): a sheet's extent counts them. */
  bare: Map<number, Set<number>>
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
export function setHidden(ctx: RuntimeContext, s: number, sheet: SheetInput) {
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

/** Sets a cell's input (null removes it), and whether it is one without a value or formula. */
export function setInput(state: EngineState, s: number, addr: string, r: number, c: number, input: CellInput | null) {
  const cells = state.input.sheets[s].cells
  if (input === null) delete cells[addr]
  else cells[addr] = input
  markBare(state, s, r, c, input)
}

function emptySheet(name: string): SheetData {
  return { name, cols: [], rows: 0, fills: [] }
}


export function invalidError(code: CheckCode): ErrorCode {
  if (code === 'unknown_function' || code === 'unknown_name') return 'NAME'
  if (code === 'unknown_sheet' || code === 'external_ref') return 'REF'
  return 'ERROR'
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
  ...(area ? { area } : {}),
})

/** A formula cell compiled: its code, and the area of an old-kind array formula. */
function compiled(state: EngineState, id: number, s: number, row: number, col: number, cell: CellInput, p: Extract<Prepared, { ok: true }>): Formula {
  const area = cell.array === undefined ? undefined : (arrayArea(formatAddress(row, col), cell.array) ?? undefined)
  const g = state.compiler.formula(codegenContext(state, s, area), p.ast, s, row, col)
  return { id, s, r: row, c: col, text: p.text, ...g, ...(area ? { area } : {}) }
}

type Prepared = { ok: true; text: string; ast: Ast } | { ok: false; invalid: Invalid }

function prepare(state: EngineState, s: number, row: number, col: number, text: string): Prepared {
  const named = (name: string, sheet: string | undefined) => nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)
  const known = (name: string, sheet: string | undefined) => !!named(name, sheet)
  const enclosing = tableAt(state.tables, s, row, col)?.name
  const r = validate(text, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables: tableNames(state.tables), ...(enclosing ? { table: enclosing } : {}) })
  if (!r.ok) return { ok: false, invalid: { code: r.code, message: r.message, at: r.at, error: invalidError(r.code) } }
  // sheets and names spelled as the book has them, as Excel shows them (`ставка*2` is `Ставка*2`)
  const sheetOf = (sh: string) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!]
  return { ok: true, text: printFormula(r.ast, { sheetOf, nameOf: (name, sheet) => named(name, sheet)?.name ?? name }), ast: r.ast }
}

// ---------- assembly ----------

/** A book compiled and computed at once. */
export const createState = (book: BookInput, opts: EngineOptions): EngineState => run(buildState(book, opts))

/** A book compiled and computed, in steps: between two formulas compiled or computed. */
export function* buildState(book: BookInput, opts: EngineOptions): Steps<EngineState> {
  const sheetNames = book.sheets.map((s) => s.name)
  const sheetIdx = new Map(sheetNames.map((n, i) => [sheetKey(n), i]))
  if (sheetIdx.size !== sheetNames.length) throw new Error('duplicate sheet name')
  // the state's own copy of the cells (edits change it); the layout and the styles are shared
  const input: BookInput = { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: { ...s.cells } })) }
  const externalFunctions = (opts.externals?.functions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f])
  const tables = tablesOf(input)
  const { names, issues } = prepareNames(sheetIdx, input.names ?? [], externalFunctions, tableNames(tables), localeBooleans(book.locale))
  const epoch = epochOf(book.dateSystem)
  // the clock's year completes date text without a year (ru-RU `1.5`); reading it marks the formula being computed
  const clock = markingClock(() => ymd(Math.floor(ctx.now) + epoch).y)
  const cache = opts.cache ?? heapCaches
  const ctx: RuntimeContext = {
    sheets: sheetNames.map(emptySheet),
    now: opts.now - epoch,
    seed: opts.seed,
    rules: rulesOf(book.locale, clock, epoch),
    externals: opts.externals?.values ?? {},
    cache,
    flat: new RangeCache(cache),
    index: new RangeCache(cache),
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
    nameIssues: issues,
    tables,
    externalFunctions,
    ctx,
    rt,
    compiler: makeCompiler(rt, opts.execution ?? 'closures'),
    formulas: [],
    at: new Map(),
    invalid: new Map(),
    succ: [],
    spillSucc: new Map(),
    topo: [],
    order: [],
    listeners: new AreaIndex(),
    volatile: new Set(),
    clocked: new Set(),
    iterated: [],
    spills: newSpills(),
    bare: new Map(),
  }
  const pending: Formula[] = []
  for (let s = 0; s < input.sheets.length; s++) {
    const sheet = input.sheets[s]
    for (const [addr, cellIn] of Object.entries(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      if (holds(cellIn)) state.spills.occupied.set(s, row, col, true, true)
      if (cellIn.formula === undefined) {
        const v = bookScalar(cellIn.value, epoch)
        if (v !== null) put(ctx.sheets[s], row, col, v)
        else markBare(state, s, row, col, cellIn)
        continue
      }
      if (due()) yield
      step()
      count(pending.length + state.invalid.size + 1)
      const p = prepare(state, s, row, col, cellIn.formula)
      if (!p.ok) {
        state.invalid.set(key(s, row, col), p.invalid)
        put(ctx.sheets[s], row, col, E[p.invalid.error])
        continue
      }
      const f = compiled(state, pending.length, s, row, col, cellIn, p)
      pending.push(f)
      if (f.nested) mark(ctx, s, row, col, true)
    }
  }
  state.compiler.link(pending)
  state.formulas = pending
  for (const f of state.formulas) state.at.set(key(f!.s, f!.r, f!.c), f!.id)
  rebuildGraph(state)
  yield* evaluateAll(state)
  return state
}

/**
 * Edits applied atomically, one recomputation of dependents; returns the changed cells (`s:r:c`).
 * `withVolatile`: the clock has changed (live session), volatile cells are recomputed in the same pass.
 */
export function* applyEdits(state: EngineState, edits: CellReplace[], withVolatile = false): Steps<string[]> {
  const track = tracker()
  const dirty = new Set<number>()
  const fresh: Formula[] = []
  // the edited cells and the ranges of arrays whose formula is gone
  const touched: Area[] = []
  let structural = false
  for (const e of edits) {
    const s = state.sheetIdx.get(sheetKey(e.sheet))
    if (s === undefined) throw new Error(`unknown sheet ${e.sheet}`)
    const { row, col } = parseAddress(e.cell)
    const k = key(s, row, col)
    setInput(state, s, e.cell, row, col, e.input)
    state.spills.occupied.set(s, row, col, holds(e.input ?? undefined))
    const old = state.at.get(k)
    if (old !== undefined) {
      if (state.formulas[old]!.nested) mark(state.ctx, s, row, col, false)
      endSpill(state, state.formulas[old]!, track, false)
      state.formulas[old] = undefined
      state.clocked.delete(old)
      state.at.delete(k)
      structural = true
    }
    state.invalid.delete(k)
    touched.push({ s, r1: row, c1: col, r2: row, c2: col })
    const input: CellInput | null = e.input
    if (input?.formula !== undefined) {
      const p = prepare(state, s, row, col, input.formula)
      if (!p.ok) {
        state.invalid.set(k, p.invalid)
        write(state, s, row, col, E[p.invalid.error], track)
        continue
      }
      const f = compiled(state, state.formulas.length, s, row, col, input, p)
      if (f.nested) mark(state.ctx, s, row, col, true)
      state.formulas.push(f)
      state.at.set(k, f.id)
      fresh.push(f)
      structural = true
    } else write(state, s, row, col, input ? bookScalar(input.value, state.ctx.rules.epoch) : null, track)
  }
  count(state.at.size + state.invalid.size)
  state.compiler.link(fresh)
  if (structural) rebuildGraph(state)
  // formulas reading the cells, arrays spilled over them or waiting for them
  for (const x of [...touched, ...state.spills.touched.splice(0)]) concerned(state, x, dirty)
  for (const f of fresh) dirty.add(f.id)
  if (withVolatile) for (const id of state.volatile) dirty.add(id)
  // cells that were in a cycle are recomputed in case the cycle is broken
  if (structural) for (const f of state.formulas) if (f && isErr(state.ctx.sheets[f.s].cols[f.c]?.[f.r]) && (state.ctx.sheets[f.s].cols[f.c]![f.r] as CErr).code === 'CYCLE') dirty.add(f.id)
  yield* recompute(state, dirty, track)
  return changed(state, track)
}

/** The clock's moment (a serial number of the 1900 date system) in the book's date system. */
export function setNow(state: EngineState, now: number, seed: number) {
  state.ctx.now = now - state.ctx.rules.epoch
  state.ctx.seed = seed
}

/** A new clock: recomputes volatile cells and their dependents. */
export function* applyClock(state: EngineState, now: number, seed: number): Steps<string[]> {
  setNow(state, now, seed)
  const track = tracker()
  yield* recompute(state, new Set(state.volatile), track)
  return changed(state, track)
}

/** Merged areas of a sheet were added or removed (`rects`): the arrays whose range meets them are computed again. */
export function* applyMerges(state: EngineState, s: number, rects: Rect[]): Steps<string[]> {
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
export function* applyHidden(state: EngineState, s: number): Steps<string[]> {
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
    lines.push(`// ${quoteSheet(state.sheetNames[f.s])}!${formatAddress(f.r, f.c)} = ${f.text}${state.topo[id] < 0 ? '  (cycle)' : ''}`, f.source!)
  }
  return lines.join('\n')
}
