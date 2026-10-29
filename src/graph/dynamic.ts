/**
 * References a formula finds as it computes (INDIRECT, OFFSET, the range operator in such a formula): `Formula.reads`.
 * They join the graph as listeners and as edges from the formulas in them, kept apart from the static ones so that a
 * formula whose reads change is moved alone where the order allows it, and the graph is built again where it does not.
 * Within one settling a formula's reads only grow, so reads that jump between targets end as a cycle, not in a loop.
 */
import type { EngineState, Formula } from '../engine.js'
import { placeOf } from '../model.js'
import type { Area } from '../spill/state.js'
import { meets } from '../structure/area.js'
import type { Rng } from '../values.js'

type Listener = Area & { id: number }

export interface Dynamic {
  /** The listener entries of each formula's reads, and the formulas in them, by formula. */
  entries: Map<number, { listeners: Listener[]; from: readonly number[] }>
  /** The formulas that read a formula's cell or spill by their reads, by that formula: successors besides `succ`. */
  succ: Map<number, number[]>
  /** Formulas whose reads changed where the order could not take them in place: the graph is built again after the round. */
  moved: number[]
  /** Formulas computed in this settling: their reads grow until it ends. */
  seen: Set<number>
}

/** The most cells of reads whose formulas are found cell by cell; larger ones build the graph again. */
const CELLS = 4096

export const newDynamic = (): Dynamic => ({ entries: new Map(), succ: new Map(), moved: [], seen: new Set() })

const sameRng = (a: Rng, b: Rng) => a.s === b.s && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2

/** The reads without repeats, in the order first made. */
function distinct(list: readonly Rng[]): Rng[] {
  const seen = new Set<string>()
  return list.filter((g) => !seen.has(g.key) && !!seen.add(g.key))
}

const sameReads = (a: readonly Rng[] | undefined, b: readonly Rng[]) => !!a && a.length === b.length && a.every((g, i) => sameRng(g, b[i]))

/** The formulas whose cells or spills meet the reads, found cell by cell; undefined where they hold too many cells. */
function targets(state: EngineState, reads: readonly Rng[]): number[] | undefined {
  const out = new Set<number>()
  for (const g of reads) {
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > CELLS) return undefined
    for (let r = g.r1; r <= g.r2; r++)
      for (let c = g.c1; c <= g.c2; c++) {
        const id = state.at.get(placeOf(g.s, r, c))
        if (id !== undefined) out.add(id)
      }
    for (const [id, x] of state.spills.claims) if (x.s === g.s && meets(x, g)) out.add(id)
  }
  return [...out]
}

/** The formulas in every read of every formula, by a scan of the formulas' cells: for a read too large to go cell by cell. */
function scanned(state: EngineState, reads: readonly Rng[]): number[] {
  const out = new Set<number>()
  for (const f of state.formulas)
    if (f && reads.some((g) => g.s === f.s && f.r >= g.r1 && f.r <= g.r2 && f.c >= g.c1 && f.c <= g.c2)) out.add(f.id)
  for (const [id, x] of state.spills.claims) if (reads.some((g) => x.s === g.s && meets(x, g))) out.add(id)
  return [...out]
}

function link(state: EngineState, d: Dynamic, f: Formula, reads: readonly Rng[], from: readonly number[]) {
  const listeners = reads.map((g): Listener => ({ s: g.s, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, id: f.id }))
  for (const x of listeners) state.listeners.add(x)
  d.entries.set(f.id, { listeners, from })
  for (const id of from) {
    const list = d.succ.get(id)
    if (list) list.push(f.id)
    else d.succ.set(id, [f.id])
  }
}

function unlink(state: EngineState, d: Dynamic, f: Formula) {
  const e = d.entries.get(f.id)
  if (!e) return
  for (const x of e.listeners) state.listeners.delete(x)
  d.entries.delete(f.id)
  for (const id of e.from) {
    const list = d.succ.get(id)!
    list.splice(list.indexOf(f.id), 1)
    if (!list.length) d.succ.delete(id)
  }
}

/** The reads of every formula in the graph being built (`rebuildGraph`): listeners, and edges from the formulas in them. */
export function linkAll(state: EngineState) {
  const d = state.dyn
  if (!d) return
  d.entries.clear()
  d.succ.clear()
  for (const f of state.formulas) if (f?.reads?.length) link(state, d, f, f.reads, targets(state, f.reads) ?? scanned(state, f.reads))
}

/**
 * A dynamic formula computed with these reads: kept, and the graph told where they changed. In place when every formula
 * in them comes before this one in the order (it read their values computed), else after the round (`moved`).
 */
export function readsMade(state: EngineState, f: Formula, made: readonly Rng[]) {
  const d = state.dyn!
  const grown = d.seen.has(f.id) && f.reads ? [...f.reads, ...made] : made
  d.seen.add(f.id)
  const reads = distinct(grown)
  if (sameReads(f.reads, reads)) return
  f.reads = reads
  const from = targets(state, reads)
  const at = state.topo[f.id]
  if (from && at >= 0 && from.every((id) => id !== f.id && state.topo[id] >= 0 && state.topo[id] < at)) {
    unlink(state, d, f)
    link(state, d, f, reads, from)
  } else d.moved.push(f.id)
}
