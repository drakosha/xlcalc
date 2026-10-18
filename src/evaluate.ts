/**
 * Evaluation: formulas computed in topological order, their values written into the sheets, with the changed
 * cells tracked. An array spills (`spill/place.ts`); when that changes cells already read, or frees cells an
 * array waits for, the formulas concerned are computed again in another round, until nothing changes.
 */
import { isMany } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import type { EngineState, Formula } from './engine.js'
import { due, located, paced, step, type Steps } from './exec/budget.js'
import { formatAddress, placeOf, placeParts } from './model.js'
import { markVolatile, readersOf, rebuildGraph, successors } from './graph.js'
import type { RuntimeContext, SheetData } from './runtime.js'
import { Column } from './column.js'
import { heldAt, unhold } from './held.js'
import { describeAt, endSpill, fixed, release, spill } from './spill/place.js'
import type { Area } from './spill/state.js'
import { meets } from './structure/area.js'
import { sheetKey } from './syntax.js'
import { E, isErr, Rng, type V } from './values.js'

export function same(a: V, b: V): boolean {
  // two numbers, the most frequent pair, compare without looking for errors (NaN is the same as NaN, as Object.is has it)
  if (typeof a === 'number' && typeof b === 'number') return a === b || (a !== a && b !== b)
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code
  return Object.is(a, b) || a === b
}

/** A value as a cell shows it: blank is 0 (a formula over an empty cell, an empty element of an array), not finite is #NUM!. */
export const shown = (v: V): V => (v === null || v === undefined ? 0 : typeof v === 'number' && !Number.isFinite(v) ? E.NUM : v)

/** Drops the cached large ranges and indexes that meet an area of a sheet. */
export function invalidate(ctx: RuntimeContext, x: Area) {
  ctx.flat.invalidate(x)
  ctx.index.invalidate(x)
  ctx.lower.invalidate(x)
  ctx.selections.invalidate(x)
  ctx.lookup?.invalidate(x)
  ctx.aggregates.invalidate(x)
}

/** The values cells had before a change, and the spills of formulas (by their cell) as `describeAt` gives them, by place (`placeOf`). */
export interface Tracker {
  cells: Map<number, V>
  spills: Map<number, string>
}

export const tracker = (): Tracker => ({ cells: new Map(), spills: new Map() })

/** Records a cell's value before it first changes. */
export function track(state: EngineState, t: Tracker | undefined, s: number, r: number, c: number) {
  if (!t) return
  const k = placeOf(s, r, c)
  if (!t.cells.has(k)) t.cells.set(k, state.rt.cell(s, r, c))
}

export function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= new Column()).set(r, v)
  if (r + 1 > sh.rows) sh.rows = r + 1
}

export function write(state: EngineState, s: number, r: number, c: number, v: V, t?: Tracker): boolean {
  const sh = state.ctx.sheets[s]
  if (same(sh.cols[c]?.get(r) ?? null, v)) return false
  track(state, t, s, r, c)
  if (state.capture) track(state, state.capture, s, r, c)
  put(sh, r, c, v)
  invalidate(state.ctx, { s, r1: r, c1: c, r2: r, c2: c })
  return true
}

/**
 * Computes the formulas `ids` in turn and writes their values; a formula in a cycle (or over one) holds its value
 * (`held.ts`), unless `iterating`. Whether any value changed as Excel's iteration counts a change (the workbook set
 * `iterate`): a number by `delta` or more, a number turned into another type or back; text, a boolean or an error
 * changed into another of them does not count. Between two formulas the steps yield when the computation's slice is
 * used up.
 */
function* evaluate(state: EngineState, ids: number[], t?: Tracker, iterating = false, delta = 0, active?: Set<number>): Steps<boolean> {
  const { rt, spills } = state
  const independent = active && active.size > 1
  if (ids.length >= 64) state.ctx.aggregates.invalidateCells(ids.map((id) => state.formulas[id]!))
  // a formula that reads the clock's year joins the volatile ones; one that no longer does leaves them
  const { clock } = state.ctx.rules
  let clocked = false
  let moved = false
  for (const id of ids) {
    if (active && !active.has(id)) {
      if (paced(1)) yield
      continue
    }
    if (due()) yield
    step()
    const f = state.formulas[id]!
    if (id >= spills.done.length) spills.done = grown(spills.done, state.formulas.length)
    spills.done[id] = ++spills.step
    if (clock) clock.read = false
    let v: V, array = false
    const was = state.ctx.sheets[f.s].cols[f.c]?.get(f.r) ?? null
    if (state.topo[id] < 0 && !iterating) {
      endSpill(state, f, t, true)
      v = heldAt(state, f, shown(was))
    } else {
      // out of a cycle's way, a formula holding a value drops it once the computation settles (a spill may close the cycle again)
      if (state.held.size && state.held.has(id)) state.outside.add(id)
      try {
        const x = f.fn()
        if (f.area) { array = true; v = fixed(state, f, x, t) }
        else if (isMany(x)) { array = true; v = spill(state, f, x instanceof Arr ? x : (rt.view(x) as Arr), t) }
        else {
          endSpill(state, f, t, false)
          v = shown(rt.scalar(x))
        }
      } catch (e) {
        throw located(e, () => ({ sheet: state.sheetNames[f.s], cell: formatAddress(f.r, f.c) }))
      }
    }
    if (clock && clock.read !== state.clocked.has(id)) {
      if (clock.read) state.clocked.add(id)
      else state.clocked.delete(id)
      clocked = true
    }
    moved ||= typeof v === 'number' && typeof was === 'number' ? Math.abs(v - was) >= delta : (typeof v === 'number') !== (typeof was === 'number')
    const changed = write(state, f.s, f.r, f.c, v, t)
    // One changed root takes the full pass; independent roots activate only the paths whose values changed.
    if (active && (array || (changed && !independent))) active = undefined
    else if (active && changed) {
      const stack = [id]
      while (stack.length) {
        for (const next of successors(state, stack.pop()!)) if (!active.has(next)) {
          active.add(next)
          if (!state.formulas[next]) stack.push(next)
        }
      }
    }
    state.ctx.cache.fit()
  }
  if (clocked) markVolatile(state)
  return moved
}

/** Formulas in a cycle or over one (computed after the others). */
const cyclic = (state: EngineState) => state.formulas.filter((f): f is Formula => !!f && state.topo[f.id] < 0).map((f) => f.id)

/**
 * Iterative calculation (`BookInput.iteration`, the workbook set `iterate`), as Excel computes: every formula in a
 * cycle or over one, sheet by sheet in the order of the sheets' names (in any case) and row by row, from
 * 0, again and again up to the iteration count, until no value changes by the maximum change or more. Each
 * computation starts from 0, so it does not depend on the values a session held before.
 */
function* iterate(state: EngineState, t?: Tracker): Steps<void> {
  const { maxIterations = 100, maxChange = 0.001 } = state.input.iteration!
  const name = (s: number) => sheetKey(state.sheetNames[s])
  const ids = cyclic(state).sort((a, b) => {
    const [f, g] = [state.formulas[a]!, state.formulas[b]!]
    return f.s !== g.s ? (name(f.s) < name(g.s) ? -1 : 1) : f.r - g.r || f.c - g.c
  })
  for (const id of ids) {
    const f = state.formulas[id]!
    release(state, id, t)
    write(state, f.s, f.r, f.c, 0, t)
  }
  state.iterated = ids
  for (let k = 0; k < maxIterations && (yield* evaluate(state, ids, t, true, maxChange)); k++);
}

/**
 * The cycles share their iterations, so a cycle added or gone changes the others: when the formulas in cycles are not
 * those the last iterative calculation computed, all of them, then and now, go to `out`.
 */
function cyclesMoved(state: EngineState, out: Set<number>) {
  const now = cyclic(state)
  const was = new Set(state.iterated)
  if (now.length === was.size && now.every((id) => was.has(id))) return
  for (const id of [...now, ...was]) if (state.formulas[id]) out.add(id)
  state.iterated = now
}

export function* evaluateAll(state: EngineState, t?: Tracker): Steps<void> {
  yield* settle(state, [state.order, cyclic(state)], t)
}

/**
 * Recomputes a set of formulas and everything that depends on them, in topological order. Where a cycle runs
 * through a spill range, before or after, everything is computed again from no spills, as a new state is: what
 * such a cycle holds depends on which arrays spilled first.
 */
export function* recompute(state: EngineState, dirty: Set<number>, t: Tracker, prune = false): Steps<void> {
  if (state.input.iteration) cyclesMoved(state, dirty)
  const before = spillCycle(state)
  // Reuse the reachability set as the active set. Large root batches and existing arrays keep the full pass.
  const roots = prune && dirty.size <= 64 && !state.input.iteration && !state.spills.claims.size && !state.spills.wait.size ? [...dirty] : undefined
  const lists = ordered(state, dirty)
  let active: Set<number> | undefined
  if (roots && !lists[1].length && lists[0].length >= 64) {
    dirty.clear()
    for (const id of roots) dirty.add(id)
    active = dirty
  }
  yield* settle(state, lists, t, active)
  if (!before && !spillCycle(state)) return
  for (const id of [...state.spills.on.keys()]) release(state, id, t)
  for (const m of [state.spills.claims, state.spills.wait, state.spills.why]) m.clear()
  rebuildGraph(state)
  yield* evaluateAll(state, t)
}

/** A formula in a cycle keeps the range it spilled into (its claim): a cycle runs through a spill range. */
const spillCycle = (state: EngineState) => [...state.spills.claims.keys()].some((id) => state.topo[id] < 0)

/**
 * The formulas of `dirty` and what depends on them: in topological order, then those in a cycle in the order the walk
 * reached them (`descendants` would list them so). The walk marks the vertices it reaches (`EngineState.marks`).
 */
function ordered(state: EngineState, dirty: Set<number>): [number[], number[]] {
  const size = Math.max(state.succ.length, state.formulas.length)
  if (state.marks.length < size) state.marks = new Uint32Array(Math.max(size, state.marks.length * 2))
  if (state.mark === 0xffffffff) state.marks.fill(state.mark = 0)
  const { marks, formulas, topo } = state
  const mark = ++state.mark
  const live: number[] = [], cycle: number[] = [], stack: number[] = []
  const reach = (id: number) => {
    marks[id] = mark
    stack.push(id)
    if (formulas[id]) (topo[id] >= 0 ? live : cycle).push(id)
  }
  for (const id of dirty) if (marks[id] !== mark) reach(id)
  while (stack.length) for (const next of successors(state, stack.pop()!)) if (marks[next] !== mark) reach(next)
  // a large part of the book: the order filtered, which costs less than sorting so many by it
  if (live.length * 16 >= state.order.length) return [state.order.filter((id) => marks[id] === mark && formulas[id]), cycle]
  return [live.sort((a, b) => topo[a] - topo[b]), cycle]
}

/** A guard on rounds: arrays settle in a few (a spill only gives way to one of lower rank, `spill/rank.ts`). */
const ROUNDS = 100_000

/** Rounds of evaluation until no spill changes cells already read or frees cells an array waits for. */
function* settle(state: EngineState, lists: [number[], number[]], t?: Tracker, active?: Set<number>): Steps<void> {
  const sp = state.spills
  for (let round = 0; ; round++) {
    if (round === ROUNDS) throw new Error('spilled arrays did not settle')
    sp.touched = []
    sp.since = sp.step
    sp.claimsChanged = false
    yield* evaluate(state, lists[0], t, false, 0, active)
    // iterative calculation computes every cycle again when any is to be computed: the cycles share their iterations
    if (state.input.iteration && lists[1].length) yield* iterate(state, t)
    else yield* evaluate(state, lists[1], t)
    const next = new Set<number>()
    for (const x of sp.touched) concerned(state, x, next, x.step)
    if (sp.claimsChanged) {
      const was = state.topo
      rebuildGraph(state)
      // formulas that went into a cycle or out of one
      if (state.input.iteration) cyclesMoved(state, next)
      else for (const f of state.formulas) if (f && (state.topo[f.id] < 0) !== (was[f.id] < 0)) next.add(f.id)
    }
    if (!next.size) {
      sp.touched.length = 0
      unhold(state)
      return
    }
    lists = ordered(state, next)
    active = undefined
  }
}

/**
 * Formulas a change of an area's cells concerns: those that read them, arrays waiting for them, arrays spilled
 * over them (a cell typed there stops them). With `step`, only those computed before that step of the round.
 */
export function concerned(state: EngineState, x: Area, out: Set<number>, step = Infinity) {
  const sp = state.spills
  const add = (id: number) => (step === Infinity || (sp.done[id] > sp.since ? sp.done[id] : -1) < step) && out.add(id)
  // without a step every reader goes, in the order the readers are found
  if (step === Infinity) readersOf(state, x, out)
  else {
    const readers = new Set<number>()
    readersOf(state, x, readers)
    readers.forEach(add)
  }
  for (const m of [sp.wait, sp.on]) for (const [id, a] of m) if (a.s === x.s && meets(a, x)) add(id)
}

/** A formula's steps in room for `n` formulas, at least twice as many as before. */
function grown(done: Float64Array, n: number): Float64Array {
  const out = new Float64Array(Math.max(n, done.length * 2, 64))
  out.set(done)
  return out
}

/** The places (`placeOf`) of the cells whose value or spill differs from the one tracked, in the order they first changed. */
export function changed(state: EngineState, t: Tracker): number[] {
  const out: number[] = []
  for (const [k, old] of t.cells) {
    const [s, r, c] = placeParts(k)
    if (!same(old, state.rt.cell(s, r, c))) out.push(k)
  }
  if (t.spills.size) {
    const seen = new Set(out)
    for (const [k, old] of t.spills) if (old !== describeAt(state, k) && !seen.has(k)) out.push(k)
  }
  return out
}
