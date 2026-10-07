/**
 * Evaluation: formulas computed in topological order, their values written into the sheets, with the changed
 * cells tracked. An array spills (`spill/place.ts`); when that changes cells already read, or frees cells an
 * array waits for, the formulas concerned are computed again in another round, until nothing changes.
 */
import { isMany } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import { key, type EngineState, type Formula } from './engine.js'
import { markVolatile, readersOf, rebuildGraph, successors } from './graph.js'
import type { RuntimeContext, SheetData } from './runtime.js'
import { describeAt, endSpill, release, spill } from './spill/place.js'
import type { Area } from './spill/state.js'
import { meets } from './structure/area.js'
import { E, isErr, Rng, type V } from './values.js'

export function same(a: V, b: V): boolean {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code
  return Object.is(a, b) || a === b
}

/** A value as a cell shows it: blank is 0 (a formula over an empty cell, an empty element of an array), not finite is #NUM!. */
export const shown = (v: V): V => (v === null || v === undefined ? 0 : typeof v === 'number' && !Number.isFinite(v) ? E.NUM : v)

/** Drops the cached large ranges and indexes that meet an area of a sheet. */
export function invalidate(ctx: RuntimeContext, x: Area) {
  for (const cache of [ctx.flat, ctx.index] as Map<string, unknown>[]) {
    if (cache.size === 0) continue
    for (const k of cache.keys()) {
      const [ks, r1, c1, r2, c2] = k.split(':').map(Number)
      if (ks === x.s && meets({ r1, c1, r2, c2 }, x)) cache.delete(k)
    }
  }
}

/** The values cells had before a change (by `s:r:c`), and the spills of formulas (by their cell) as `describeSpill` gives them. */
export interface Tracker {
  cells: Map<string, V>
  spills: Map<string, string>
}

export const tracker = (): Tracker => ({ cells: new Map(), spills: new Map() })

/** Records a cell's value before it first changes. */
export function track(state: EngineState, t: Tracker | undefined, s: number, r: number, c: number) {
  const k = key(s, r, c)
  if (t && !t.cells.has(k)) t.cells.set(k, state.rt.cell(s, r, c))
}

export function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= [])[r] = v
  if (r + 1 > sh.rows) sh.rows = r + 1
}

export function write(state: EngineState, s: number, r: number, c: number, v: V, t?: Tracker): boolean {
  const sh = state.ctx.sheets[s]
  if (same(sh.cols[c]?.[r] ?? null, v)) return false
  track(state, t, s, r, c)
  put(sh, r, c, v)
  invalidate(state.ctx, { s, r1: r, c1: c, r2: r, c2: c })
  return true
}

function evaluate(state: EngineState, ids: number[], t?: Tracker) {
  const { rt, spills } = state
  // a formula that reads the clock's year joins the volatile ones; one that no longer does leaves them
  const { clock } = state.ctx.rules
  let clocked = false
  for (const id of ids) {
    const f = state.formulas[id]!
    spills.done.set(id, ++spills.step)
    if (clock) clock.read = false
    let v: V
    if (state.topo[id] < 0) {
      endSpill(state, f, t, true)
      v = E.CYCLE
    } else {
      const x = f.fn()
      if (isMany(x)) v = spill(state, f, x instanceof Arr ? x : (rt.view(x) as Arr), t)
      else {
        endSpill(state, f, t, false)
        v = shown(rt.scalar(x))
      }
    }
    if (clock && clock.read !== state.clocked.has(id)) {
      if (clock.read) state.clocked.add(id)
      else state.clocked.delete(id)
      clocked = true
    }
    write(state, f.s, f.r, f.c, v, t)
  }
  if (clocked) markVolatile(state)
}

/** Formulas in a cycle (computed after the others). */
const cyclic = (state: EngineState) => state.formulas.filter((f): f is Formula => !!f && state.topo[f.id] < 0).map((f) => f.id)

export function evaluateAll(state: EngineState, t?: Tracker) {
  settle(state, [state.order, cyclic(state)], t)
}

/**
 * Recomputes a set of formulas and everything that depends on them, in topological order. Where a cycle runs
 * through a spill range, before or after, everything is computed again from no spills, as a new state is: what
 * such a cycle holds depends on which arrays spilled first.
 */
export function recompute(state: EngineState, dirty: Set<number>, t: Tracker) {
  const before = spillCycle(state)
  settle(state, ordered(state, dirty), t)
  if (!before && !spillCycle(state)) return
  for (const id of [...state.spills.on.keys()]) release(state, id, t)
  for (const m of [state.spills.claims, state.spills.wait, state.spills.why]) m.clear()
  rebuildGraph(state)
  evaluateAll(state, t)
}

/** A formula in a cycle keeps the range it spilled into (its claim): a cycle runs through a spill range. */
const spillCycle = (state: EngineState) => [...state.spills.claims.keys()].some((id) => state.topo[id] < 0)

/** The formulas of `dirty` and what depends on them: in topological order, then those in a cycle. */
function ordered(state: EngineState, dirty: Set<number>): [number[], number[]] {
  const stack = [...dirty]
  while (stack.length)
    for (const n of successors(state, stack.pop()!))
      if (!dirty.has(n)) {
        dirty.add(n)
        stack.push(n)
      }
  const live = [...dirty].filter((id) => state.formulas[id])
  return [live.filter((id) => state.topo[id] >= 0).sort((a, b) => state.topo[a] - state.topo[b]), live.filter((id) => state.topo[id] < 0)]
}

/** A guard on rounds: arrays settle in a few (a spill only gives way to one of lower rank, `spill/rank.ts`). */
const ROUNDS = 100_000

/** Rounds of evaluation until no spill changes cells already read or frees cells an array waits for. */
function settle(state: EngineState, lists: [number[], number[]], t?: Tracker) {
  const sp = state.spills
  for (let round = 0; ; round++) {
    if (round === ROUNDS) throw new Error('spilled arrays did not settle')
    sp.touched = []
    sp.done.clear()
    sp.claimsChanged = false
    evaluate(state, lists[0], t)
    evaluate(state, lists[1], t)
    const next = new Set<number>()
    for (const x of sp.touched) concerned(state, x, next, x.step)
    if (sp.claimsChanged) {
      rebuildGraph(state)
      // formulas that went into a cycle or out of one
      for (const f of state.formulas) if (f && (state.topo[f.id] < 0) !== isCycle(state.rt.cell(f.s, f.r, f.c))) next.add(f.id)
    }
    if (!next.size) return
    lists = ordered(state, next)
  }
}

const isCycle = (v: V) => isErr(v) && v.code === 'CYCLE'

/**
 * Formulas a change of an area's cells concerns: those that read them, arrays waiting for them, arrays spilled
 * over them (a cell typed there stops them). With `step`, only those computed before that step of the round.
 */
export function concerned(state: EngineState, x: Area, out: Set<number>, step = Infinity) {
  const sp = state.spills
  const add = (id: number) => (sp.done.get(id) ?? -1) < step && out.add(id)
  const readers = new Set<number>()
  readersOf(state, x, readers)
  readers.forEach(add)
  for (const m of [sp.wait, sp.on]) for (const [id, a] of m) if (a.s === x.s && meets(a, x)) add(id)
}

export function changed(state: EngineState, t: Tracker): string[] {
  const out = new Set<string>()
  for (const [k, old] of t.cells) {
    const [s, r, c] = k.split(':').map(Number)
    if (!same(old, state.rt.cell(s, r, c))) out.add(k)
  }
  for (const [k, old] of t.spills) if (old !== describeAt(state, k)) out.add(k)
  return [...out]
}
