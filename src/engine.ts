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
import { makeCompiler, PRELUDE, type CodegenContext, type Compiler, type FormulaCode } from './compile.js'
import { changed, concerned, evaluateAll, put, recompute, tracker, write } from './evaluate.js'
import { FUNCTIONS } from './functions.js'
import { rebuildGraph } from './graph.js'
import { ymd } from './dates.js'
import { rulesOf } from './locale/rules.js'
import type { RulesClock } from './locale/rules/data.js'
import { formatAddress, inputScalar, MAX_COLS, parseAddress } from './model.js'
import { makeRuntime, type Runtime, type RuntimeContext, type SheetData } from './runtime.js'
import { nameFor, prepareNames, type PreparedName } from './names.js'
import { blocks } from './spill/occupied.js'
import { endSpill, spillRef } from './spill/place.js'
import { newSpills, type Area, type Spills } from './spill/state.js'
import { meets, type Rect } from './structure/area.js'
import { canonical, quoteSheet, sheetKey, validate, type Ast } from './syntax.js'
import { CErr, E, isErr, type V } from './values.js'

export interface Formula extends FormulaCode {
  id: number
  s: number
  r: number
  c: number
  /** Canonical text. */
  text: string
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
  /** The formulas reading a spilled array's range, by the array's formula (successors besides `succ`). */
  spillSucc: Map<number, number[]>
  /** Position in the topological order; -1 for cells in a cycle and cells depending on them. */
  topo: number[]
  order: number[]
  /** Listeners of narrow ranges: sheet → column → [r1, r2, id]. */
  listeners: Map<number, Map<number, [number, number, number][]>>
  /** Listeners of wide ranges (whole rows and more than 64 columns): sheet → [r1, r2, c1, c2, id]. */
  wide: Map<number, [number, number, number, number, number][]>
  /** Formulas that depend on volatile functions or on `clocked` formulas (directly or through others). */
  volatile: Set<number>
  /** Formulas whose latest computation read the clock's year (date text without a year, by the book's locale). */
  clocked: Set<number>
  spills: Spills
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
  return { name, cols: [], rows: 0, fills: [] }
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
  const r = validate(text, { sheets: state.sheetNames, names: visible, externalFunctions: state.externalFunctions })
  if (!r.ok) return { ok: false, invalid: { code: r.code, message: r.message, at: r.at, error: invalidError(r.code) } }
  return { ok: true, text: canonical(r.ast, (sh) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!]), ast: r.ast }
}

// ---------- assembly ----------

export function createState(book: BookInput, opts: EngineOptions): EngineState {
  const sheetNames = book.sheets.map((s) => s.name)
  const sheetIdx = new Map(sheetNames.map((n, i) => [sheetKey(n), i]))
  if (sheetIdx.size !== sheetNames.length) throw new Error('duplicate sheet name')
  // the state's own copy of the cells (edits change it); the layout and the styles are shared
  const input: BookInput = { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: { ...s.cells } })) }
  const { names, issues } = prepareNames(sheetIdx, input.names ?? [])
  // the clock's year completes date text without a year (ru-RU `1.5`); reading it marks the formula being computed
  const clock: RulesClock = {
    read: false,
    year: () => {
      clock.read = true
      return ymd(Math.floor(ctx.now)).y
    },
  }
  const ctx: RuntimeContext = {
    sheets: sheetNames.map(emptySheet),
    now: opts.now,
    seed: opts.seed,
    rules: rulesOf(book.locale, clock),
    externals: opts.externals?.values ?? {},
    flat: new Map(),
    index: new Map(),
    nested: new Map(),
    spillRef: (s, r, c) => spillRef(state, s, r, c),
  }
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
    spillSucc: new Map(),
    topo: [],
    order: [],
    listeners: new Map(),
    wide: new Map(),
    volatile: new Set(),
    clocked: new Set(),
    spills: newSpills(),
  }
  const pending: Formula[] = []
  input.sheets.forEach((sheet, s) => {
    for (const [addr, cellIn] of Object.entries(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      if (blocks(cellIn)) state.spills.occupied.set(s, row, col, true, true)
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
    const cells = state.input.sheets[s].cells
    if (e.input === null) delete cells[e.cell]
    else cells[e.cell] = e.input
    state.spills.occupied.set(s, row, col, blocks(e.input ?? undefined))
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
  // formulas reading the cells, arrays spilled over them or waiting for them
  for (const x of [...touched, ...state.spills.touched.splice(0)]) concerned(state, x, dirty)
  for (const f of fresh) dirty.add(f.id)
  if (withVolatile) for (const id of state.volatile) dirty.add(id)
  // cells that were in a cycle are recomputed in case the cycle is broken
  if (structural) for (const f of state.formulas) if (f && isErr(state.ctx.sheets[f.s].cols[f.c]?.[f.r]) && (state.ctx.sheets[f.s].cols[f.c]![f.r] as CErr).code === 'CYCLE') dirty.add(f.id)
  recompute(state, dirty, track)
  return changed(state, track)
}

/** A new clock: recomputes volatile cells and their dependents. */
export function applyClock(state: EngineState, now: number, seed: number): string[] {
  state.ctx.now = now
  state.ctx.seed = seed
  const track = tracker()
  recompute(state, new Set(state.volatile), track)
  return changed(state, track)
}

/** Merged areas of a sheet were added or removed (`rects`): the arrays whose range meets them are computed again. */
export function applyMerges(state: EngineState, s: number, rects: Rect[]): string[] {
  const track = tracker()
  const dirty = new Set<number>()
  state.spills.merges.delete(s)
  for (const m of [state.spills.on, state.spills.wait]) for (const [id, x] of m) if (x.s === s && rects.some((r) => meets(r, x))) dirty.add(id)
  recompute(state, dirty, track)
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

