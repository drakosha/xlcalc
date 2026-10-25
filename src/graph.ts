/**
 * The reference graph over a book's formulas: which formulas read which cells, and the order of evaluation.
 * Besides the references, a spilled array (its claim, `spill/state.ts`) comes before the formulas reading its range.
 */
import type { EngineState, Formula } from './engine.js'
import { spend } from './exec/budget.js'
import { bySheet, pastEnd } from './exec/walk.js'
import type { Area } from './spill/state.js'
import { Rng } from './values.js'
import { placeOf } from './model.js'
import { AreaIndex } from './area-index.js'
import { RangeDependencies } from './graph/ranges.js'
import { indegrees, releaseReady } from './graph/order.js'

// a range operator reaching a spilled array reads as far as the array claims, around its other ends; LOOKUP past it
export function referencesOf(state: EngineState, f: Formula): readonly Rng[] {
  if (!f.spans) return f.refs
  const more: Rng[] = []
  for (const span of f.spans) {
    const reached = [...span.boxes]
    for (const a of span.anchors) {
      const id = state.at.get(placeOf(a.s, a.r1, a.c1))
      const x = id === undefined ? undefined : state.spills.claims.get(id)
      if (x) reached.push(new Rng(x.s, x.r1, x.c1, x.r2, x.c2))
      // LOOKUP reads on past the array in its direction, or past the cell when it has not spilled
      if (span.past) more.push(...pastEnd(x ? new Rng(x.s, x.r1, x.c1, x.r2, x.c2) : a, span.past, span.part))
    }
    if (!span.past && reached.length > span.boxes.length) more.push(...bySheet(reached).values())
  }
  return more.length ? [...f.refs, ...more] : f.refs
}

/** The reference graph of the formulas: successors, the topological order, listeners of ranges, volatile formulas. */
export function rebuildGraph(state: EngineState) {
  const { formulas } = state
  const graph = new RangeDependencies(formulas)
  const { succ } = graph
  const listeners: EngineState['listeners'] = new AreaIndex()
  for (const f of formulas) {
    if (!f) continue
    spend(1)
    for (const g of referencesOf(state, f)) {
      listeners.add({ s: g.s, r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2, id: f.id })
      graph.add(g, f.id)
    }
  }
  // each list of successors as an array of its own length: one filled by pushes keeps room for many more, for each vertex
  for (let i = 0; i < succ.length; i++) if (succ[i].length) succ[i] = succ[i].slice()
  state.succ = succ
  state.listeners = listeners
  const spillSucc = new Map<number, number[]>()
  for (const [id, x] of state.spills.claims) {
    const readers = new Set<number>()
    readersOf(state, x, readers)
    if (!readers.size) continue
    spillSucc.set(id, [...readers])
  }
  state.spillSucc = spillSucc
  state.spills.rank = state.spills.looped = state.spills.pred = undefined
  const indeg = indegrees(succ, spillSucc)
  const order: number[] = []
  const queue: number[] = []
  formulas.forEach((f, id) => f && indeg[id] === 0 && queue.push(id))
  const ready = (id: number) => queue.push(id)
  // the queue need not be ordered by address: values do not depend on the evaluation order (of two arrays
  // over one cell, the one that spills is chosen by rank, `spill/rank.ts`)
  while (queue.length) {
    const id = queue.pop()!
    order.push(id)
    releaseReady(id, succ, indeg, formulas.length, ready, spillSucc)
  }
  const topo = new Array<number>(formulas.length).fill(-1)
  order.forEach((id, i) => (topo[id] = i))
  state.topo = topo
  state.order = order
  markVolatile(state)
}

/** The formulas that depend on the clock: volatile functions and the year of dates without one (`clocked`), and the formulas over them. */
export function markVolatile(state: EngineState) {
  const roots = new Set([...state.formulas.filter((f): f is Formula => !!f?.volatile).map((f) => f.id), ...state.clocked])
  state.volatile = new Set([...descendants(state, roots)].filter((id) => state.formulas[id]))
}

/** Expand roots through formula and helper vertices; callers retain only the formula ids they need. */
function descendants(state: EngineState, seen: Set<number>): Set<number> {
  const stack = [...seen]
  while (stack.length) {
    const id = stack.pop()!
    for (const n of successors(state, id)) if (!seen.has(n)) {
      seen.add(n)
      stack.push(n)
    }
  }
  return seen
}

/** Next graph vertices, including the formulas reading a spill's range. Helper ids follow the formula slots. */
export function successors(state: EngineState, id: number): number[] {
  const spilled = state.spillSucc.get(id)
  return spilled ? [...state.succ[id], ...spilled] : (state.succ[id] ?? [])
}

/** Formulas that read a cell of an area (by their references, whether or not the cells hold anything). */
export function readersOf(state: EngineState, x: Area, out: Set<number>) {
  state.listeners.forEach(x, ({ id }) => out.add(id))
}
