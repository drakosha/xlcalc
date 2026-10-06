/** Evaluation: formulas computed in topological order, their values written into the sheets, with the changed cells tracked. */
import { isMany } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import { key, type EngineState, type Formula } from './engine.js'
import type { RuntimeContext, SheetData } from './runtime.js'
import { E, isErr, Rng, type V } from './values.js'

export function same(a: V, b: V): boolean {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code
  return Object.is(a, b) || a === b
}

/** Drops the cached large ranges and indexes that a cell touches. */
export function invalidate(ctx: RuntimeContext, s: number, r: number, c: number) {
  for (const cache of [ctx.flat, ctx.index] as Map<string, unknown>[]) {
    if (cache.size === 0) continue
    for (const k of cache.keys()) {
      const [ks, r1, c1, r2, c2] = k.split(':').map(Number)
      if (ks === s && r >= r1 && r <= r2 && c >= c1 && c <= c2) cache.delete(k)
    }
  }
}

export type Tracker = Map<string, V>

export function write(state: EngineState, s: number, r: number, c: number, v: V, track?: Tracker): boolean {
  const sh = state.ctx.sheets[s]
  const old = sh.cols[c]?.[r] ?? null
  if (same(old, v)) return false
  const k = key(s, r, c)
  if (track && !track.has(k)) track.set(k, old)
  put(sh, r, c, v)
  invalidate(state.ctx, s, r, c)
  return true
}

export function evaluate(state: EngineState, ids: number[], track?: Tracker) {
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

export function evaluateAll(state: EngineState, track?: Tracker) {
  evaluate(state, state.order, track)
  const cyc = state.formulas.filter((f): f is Formula => !!f && state.topo[f.id] < 0).map((f) => f.id)
  evaluate(state, cyc, track)
}


export function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= [])[r] = v
  if (r + 1 > sh.rows) sh.rows = r + 1
}

/** Recomputes a set of formulas and everything that depends on them, in topological order. */
export function recompute(state: EngineState, dirty: Set<number>, track: Tracker) {
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

export function changed(state: EngineState, track: Tracker): string[] {
  const out: string[] = []
  for (const [k, old] of track) {
    const [s, r, c] = k.split(':').map(Number)
    if (!same(old, state.ctx.sheets[s].cols[c]?.[r] ?? null)) out.push(k)
  }
  return out
}
