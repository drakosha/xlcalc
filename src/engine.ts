/**
 * The engine: formulas are compiled (`compile.ts`) to closures or, in the `eval` mode, to JS text with one
 * `new Function` per batch; cells are computed over the reference graph.
 *
 * Each formula is its own function (editing a formula in a session compiles only that formula). A full
 * computation follows the topological order; editing a value recomputes only the dependent cells.
 *
 * No platform imports, so the engine runs outside Node as well.
 */
import type { BookInput, CellEdit, CellInput, CellValue, CheckCode, ErrorCode, Execution, Externals, NameIssue } from './contract.js'
import { isMany } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import { makeCompiler, PRELUDE, type CodegenContext, type Compiler, type FormulaCode } from './compile.js'
import { FUNCTIONS } from './functions.js'
import { listenersOf, rebuildGraph } from './graph.js'
import { formatAddress, inputScalar, MAX_COLS, parseAddress, sortAddresses } from './model.js'
import { makeRuntime, type Runtime, type RuntimeContext, type SheetData } from './runtime.js'
import { nameFor, prepareNames, type PreparedName } from './names.js'
import { canonical, quoteSheet, sheetKey, validate, type Ast } from './syntax.js'
import { CErr, E, isErr, Rng, type V } from './values.js'

export interface Formula extends FormulaCode {
  id: number
  s: number
  r: number
  c: number
  /** Canonical text. */
  text: string
  /** Rows and columns of the array result of the latest computation (the cell shows #SPILL!). */
  spill?: [number, number]
}

interface Invalid {
  code: CheckCode
  message: string
  at: number
  error: ErrorCode
}

export interface EngineOptions {
  /** The clock moment, a serial number (UTC). */
  now: number
  seed: number
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
}

export interface EngineState {
  input: BookInput
  sheetNames: string[]
  sheetIdx: Map<string, number>
  names: PreparedName[]
  nameIssues: NameIssue[]
  externalFunctions: string[]
  ctx: RuntimeContext
  rt: Runtime
  compiler: Compiler
  /** id → formula; a hole after deletion. */
  formulas: (Formula | undefined)[]
  /** `s:r:c` → formula id. */
  at: Map<string, number>
  invalid: Map<string, Invalid>
  succ: number[][]
  /** Position in the topological order; -1 for cells in a cycle and cells depending on them. */
  topo: number[]
  order: number[]
  /** Listeners of narrow ranges: sheet → column → [r1, r2, id]. */
  listeners: Map<number, Map<number, [number, number, number][]>>
  /** Listeners of wide ranges (whole rows and more than 64 columns): sheet → [r1, r2, c1, c2, id]. */
  wide: Map<number, [number, number, number, number, number][]>
  /** Formulas that depend on volatile functions (directly or through others). */
  volatile: Set<number>
}

export const key = (s: number, r: number, c: number) => `${s}:${r}:${c}`

/** Marks or unmarks a cell whose formula calls SUBTOTAL or AGGREGATE. */
function mark(ctx: RuntimeContext, s: number, r: number, c: number, on: boolean) {
  let marks = ctx.nested.get(s)
  if (!marks) ctx.nested.set(s, (marks = new Set()))
  if (on) marks.add(r * MAX_COLS + c)
  else marks.delete(r * MAX_COLS + c)
}

function emptySheet(name: string): SheetData {
  return { name, cols: [], rows: 0 }
}

function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= [])[r] = v
  if (r + 1 > sh.rows) sh.rows = r + 1
}

export function invalidError(code: CheckCode): ErrorCode {
  if (code === 'unknown_function' || code === 'unknown_name') return 'NAME'
  if (code === 'unknown_sheet' || code === 'external_ref') return 'REF'
  return 'ERROR'
}

// ---------- preparation ----------

export const codegenContext = (state: EngineState, s: number): CodegenContext => ({
  sheet: (name) => (name === null ? s : state.sheetIdx.get(sheetKey(name))!),
  name: (name) => nameFor(state.names, s, name),
})

type Prepared = { ok: true; text: string; ast: Ast } | { ok: false; invalid: Invalid }

function prepare(state: EngineState, s: number, text: string): Prepared {
  const visible = state.names.filter((n) => n.scope === -1 || n.scope === s).map((n) => n.name)
  // an array result is found when computing (#SPILL!), not refused here
  const r = validate(text, { sheets: state.sheetNames, names: visible, externalFunctions: state.externalFunctions }, (name) => nameFor(state.names, s, name)?.area, false)
  if (!r.ok) return { ok: false, invalid: { code: r.code, message: r.message, at: r.at, error: invalidError(r.code) } }
  return { ok: true, text: canonical(r.ast, (sh) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!]), ast: r.ast }
}

// ---------- evaluation ----------

export function same(a: V, b: V): boolean {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code
  return Object.is(a, b) || a === b
}

/** Drops the cached large ranges and indexes that a cell touches. */
function invalidate(ctx: RuntimeContext, s: number, r: number, c: number) {
  for (const cache of [ctx.flat, ctx.index] as Map<string, unknown>[]) {
    if (cache.size === 0) continue
    for (const k of cache.keys()) {
      const [ks, r1, c1, r2, c2] = k.split(':').map(Number)
      if (ks === s && r >= r1 && r <= r2 && c >= c1 && c <= c2) cache.delete(k)
    }
  }
}

type Tracker = Map<string, V>

function write(state: EngineState, s: number, r: number, c: number, v: V, track?: Tracker): boolean {
  const sh = state.ctx.sheets[s]
  const old = sh.cols[c]?.[r] ?? null
  if (same(old, v)) return false
  const k = key(s, r, c)
  if (track && !track.has(k)) track.set(k, old)
  put(sh, r, c, v)
  invalidate(state.ctx, s, r, c)
  return true
}

function evaluate(state: EngineState, ids: number[], track?: Tracker) {
  const { rt } = state
  for (const id of ids) {
    const f = state.formulas[id]!
    let v: V
    f.spill = undefined
    if (state.topo[id] < 0) v = E.CYCLE
    else {
      const x = f.fn()
      // an array of several values cannot be shown in one cell before spilling: #SPILL!
      if (isMany(x)) {
        f.spill = x instanceof Arr ? [x.rows, x.cols] : [x.r2 - x.r1 + 1, x.c2 - x.c1 + 1]
        v = E.SPILL
      } else v = x instanceof Rng || x instanceof Arr ? rt.scalar(x) : x
      // a formula referencing an empty cell gives 0, as in Excel
      if (v === null || v === undefined) v = 0
      else if (typeof v === 'number' && !Number.isFinite(v)) v = E.NUM
    }
    write(state, f.s, f.r, f.c, v, track)
  }
}

function evaluateAll(state: EngineState, track?: Tracker) {
  evaluate(state, state.order, track)
  const cyc = state.formulas.filter((f): f is Formula => !!f && state.topo[f.id] < 0).map((f) => f.id)
  evaluate(state, cyc, track)
}

// ---------- assembly ----------

export function createState(book: BookInput, opts: EngineOptions): EngineState {
  const sheetNames = book.sheets.map((s) => s.name)
  const sheetIdx = new Map(sheetNames.map((n, i) => [sheetKey(n), i]))
  if (sheetIdx.size !== sheetNames.length) throw new Error('duplicate sheet name')
  // the state's own copy of the cells (edits change it); the layout and the styles are shared
  const input: BookInput = { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: { ...s.cells } })) }
  const { names, issues } = prepareNames(sheetIdx, input.names ?? [])
  const ctx: RuntimeContext = { sheets: sheetNames.map(emptySheet), now: opts.now, seed: opts.seed, externals: opts.externals?.values ?? {}, flat: new Map(), index: new Map(), nested: new Map() }
  const rt = makeRuntime(ctx)
  const state: EngineState = {
    input,
    sheetNames,
    sheetIdx,
    names,
    nameIssues: issues,
    externalFunctions: (opts.externals?.functions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f]),
    ctx,
    rt,
    compiler: makeCompiler(rt, opts.execution ?? 'closures'),
    formulas: [],
    at: new Map(),
    invalid: new Map(),
    succ: [],
    topo: [],
    order: [],
    listeners: new Map(),
    wide: new Map(),
    volatile: new Set(),
  }
  const pending: Formula[] = []
  input.sheets.forEach((sheet, s) => {
    for (const [addr, cellIn] of Object.entries(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      if (cellIn.formula === undefined) {
        const v = inputScalar(cellIn.value)
        if (v !== null) put(ctx.sheets[s], row, col, v)
        continue
      }
      const p = prepare(state, s, cellIn.formula)
      if (!p.ok) {
        state.invalid.set(key(s, row, col), p.invalid)
        put(ctx.sheets[s], row, col, E[p.invalid.error])
        continue
      }
      const g = state.compiler.formula(codegenContext(state, s), p.ast, s, row, col)
      pending.push({ id: pending.length, s, r: row, c: col, text: p.text, ...g })
      if (g.nested) mark(ctx, s, row, col, true)
    }
  })
  state.compiler.link(pending)
  state.formulas = pending
  for (const f of state.formulas) state.at.set(key(f!.s, f!.r, f!.c), f!.id)
  rebuildGraph(state)
  evaluateAll(state)
  return state
}

/**
 * Edits applied atomically, one recomputation of dependents; returns the changed cells (`s:r:c`).
 * `withVolatile`: the clock has changed (live session), volatile cells are recomputed in the same pass.
 */
export function applyEdits(state: EngineState, edits: CellEdit[], withVolatile = false): string[] {
  const track: Tracker = new Map()
  const dirty = new Set<number>()
  const fresh: Formula[] = []
  let structural = false
  for (const e of edits) {
    const s = state.sheetIdx.get(sheetKey(e.sheet))
    if (s === undefined) throw new Error(`unknown sheet ${e.sheet}`)
    const { row, col } = parseAddress(e.cell)
    const k = key(s, row, col)
    const cells = state.input.sheets[s].cells
    if (e.input === null) delete cells[e.cell]
    else cells[e.cell] = e.input
    const old = state.at.get(k)
    if (old !== undefined) {
      if (state.formulas[old]!.nested) mark(state.ctx, s, row, col, false)
      state.formulas[old] = undefined
      state.at.delete(k)
      structural = true
    }
    state.invalid.delete(k)
    listenersOf(state, s, row, col, dirty)
    const input: CellInput | null = e.input
    if (input?.formula !== undefined) {
      const p = prepare(state, s, input.formula)
      if (!p.ok) {
        state.invalid.set(k, p.invalid)
        write(state, s, row, col, E[p.invalid.error], track)
        continue
      }
      const g = state.compiler.formula(codegenContext(state, s), p.ast, s, row, col)
      const f: Formula = { id: state.formulas.length, s, r: row, c: col, text: p.text, ...g }
      if (g.nested) mark(state.ctx, s, row, col, true)
      state.formulas.push(f)
      state.at.set(k, f.id)
      fresh.push(f)
      structural = true
    } else write(state, s, row, col, input ? inputScalar(input.value) : null, track)
  }
  state.compiler.link(fresh)
  if (structural) rebuildGraph(state)
  for (const f of fresh) dirty.add(f.id)
  if (withVolatile) for (const id of state.volatile) dirty.add(id)
  // cells that were in a cycle are recomputed in case the cycle is broken
  if (structural) for (const f of state.formulas) if (f && isErr(state.ctx.sheets[f.s].cols[f.c]?.[f.r]) && (state.ctx.sheets[f.s].cols[f.c]![f.r] as CErr).code === 'CYCLE') dirty.add(f.id)
  recompute(state, dirty, track)
  return changed(state, track)
}

/** Recomputes a set of formulas and everything that depends on them, in topological order. */
function recompute(state: EngineState, dirty: Set<number>, track: Tracker) {
  const stack = [...dirty]
  while (stack.length) {
    const id = stack.pop()!
    for (const n of state.succ[id] ?? [])
      if (!dirty.has(n)) {
        dirty.add(n)
        stack.push(n)
      }
  }
  const live = [...dirty].filter((id) => state.formulas[id])
  const ordered = live.filter((id) => state.topo[id] >= 0).sort((a, b) => state.topo[a] - state.topo[b])
  evaluate(state, ordered, track)
  evaluate(state, live.filter((id) => state.topo[id] < 0), track)
}

function changed(state: EngineState, track: Tracker): string[] {
  const out: string[] = []
  for (const [k, old] of track) {
    const [s, r, c] = k.split(':').map(Number)
    if (!same(old, state.ctx.sheets[s].cols[c]?.[r] ?? null)) out.push(k)
  }
  return out
}

/** A new clock: recomputes volatile cells and their dependents. */
export function applyClock(state: EngineState, now: number, seed: number): string[] {
  state.ctx.now = now
  state.ctx.seed = seed
  const track: Tracker = new Map()
  recompute(state, new Set(state.volatile), track)
  return changed(state, track)
}

/** The value of a cell (null when empty). */
export function valueAt(state: EngineState, s: number, r: number, c: number): V {
  return state.ctx.sheets[s].cols[c]?.[r] ?? null
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

/** Cells of the book by address (for output and hashing). */
export function sortedCells(state: EngineState, s: number): string[] {
  const addrs = new Set(Object.keys(state.input.sheets[s].cells))
  return sortAddresses(addrs)
}
