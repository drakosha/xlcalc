/**
 * Placing a formula's array over the cells to the right of and below it (`docs/spill.md`), as Excel 365 does:
 * the range must stay on the sheet, meet no merged area and hold no value or formula but the formula's own
 * cell; of two arrays over one cell the one with the lower rank spills (`rank.ts`). An array that cannot spill
 * gives #SPILL! and waits for its range.
 */
import type { Arg } from '../args.js'
import type { Arr } from '../arrays/value.js'
import { canonValue } from '../canon.js'
import { toCellValue, key, type EngineState, type Formula } from '../engine.js'
import { allot, spend } from '../exec/budget.js'
import { invalidate, put, same, shown, track, type Tracker } from '../evaluate.js'
import { parseMerge } from '../layout/merges.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { meets, type Rect } from '../structure/area.js'
import { quoteSheet } from '../syntax.js'
import { E, Rng, type CErr, type V } from '../values.js'
import type { Fill } from './fill.js'
import { fit } from './fixed.js'
import { beats } from './rank.js'
import { equalAreas, type Area, type Spill } from './state.js'

/** An array of up to this many values is stored whole; a larger one only up to where the rest is one value. */
export const LISTED = 65_536

/** The stored part of an array from its top-left, `h` × `w`, and the one value of the rest (none when stored whole). */
function listing(a: Arr): { h: number; w: number; fill?: V } {
  if (a.rows * a.cols <= LISTED) return { h: a.rows, w: a.cols }
  const fill = shown(a.at(a.rows - 1, a.cols - 1))
  let h = 1
  let w = 1
  for (let r = 0; r < a.h; r++) {
    spend(a.w)
    for (let c = 0; c < a.w; c++)
      if (!same(shown(a.v[r * a.w + c]), fill)) {
        h = Math.max(h, r + 1)
        w = Math.max(w, c + 1)
      }
  }
  return h === a.rows && w === a.cols ? { h, w } : { h, w, fill }
}

const cellText = (state: EngineState, s: number, r: number, c: number) => `${quoteSheet(state.sheetNames[s])}!${formatAddress(r, c)}`

/** The merged areas of a sheet. */
function mergesOf(state: EngineState, s: number): Rect[] {
  let m = state.spills.merges.get(s)
  if (!m) state.spills.merges.set(s, (m = (state.input.sheets[s].merges ?? []).map((t) => parseMerge(t)!)))
  return m
}

/**
 * An array formula of the old kind fills its area with its answer fitted to it (`fixed.ts`), whatever is there: the
 * arrays spilled over it stop and wait. The formula's own value is the area's first.
 */
export function fixed(state: EngineState, f: Formula, x: Arg, t?: Tracker): V {
  const area: Area = { s: f.s, ...f.area! }
  const a = fit(x, area.r2 - area.r1 + 1, area.c2 - area.c1 + 1, state.rt.view)
  for (const o of [...state.spills.on.values()]) {
    if (o.id === f.id || o.s !== f.s || !meets(o, area)) continue
    release(state, o.id, t)
    state.spills.wait.set(o.id, { s: o.s, r1: o.r1, c1: o.c1, r2: o.r2, c2: o.c2 })
  }
  place(state, f, a, area, t)
  return shown(a.at(0, 0))
}

/** The array `a` of formula `f` spills if it can; the formula's own value (its first element, or #SPILL!). */
export function spill(state: EngineState, f: Formula, a: Arr, t?: Tracker): V {
  const sp = state.spills
  const area: Area = { s: f.s, r1: f.r, c1: f.c, r2: f.r + a.rows - 1, c2: f.c + a.cols - 1 }
  const cannot = `the array of ${a.rows}×${a.cols} values cannot spill`
  if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS) return blocked(state, f, t, `${cannot}: it would go past the edge of the sheet`)
  const merge = mergesOf(state, f.s).find((m) => meets(m, area))
  if (merge) return blocked(state, f, t, `${cannot}: it meets the merged area ${cellText(state, f.s, merge.r1, merge.c1)}:${formatAddress(merge.r2, merge.c2)}`, area)
  const cell = sp.occupied.first(f.s, area, f.r, f.c)
  if (cell) return blocked(state, f, t, `${cannot}: ${cellText(state, f.s, cell[0], cell[1])} is not empty`, area)
  const over = [...sp.on.values()].filter((x) => x.id !== f.id && x.s === f.s && meets(x, area))
  // an old-kind array formula's area is never spilled over
  const first = over.find((x) => state.formulas[x.id]?.area || !beats(state, f.id, x.id))
  if (first) return blocked(state, f, t, `${cannot}: it meets the spilled array of ${cellText(state, f.s, first.r1, first.c1)}`, area)
  // the arrays this one wins over stop spilling and wait for their range
  for (const x of over) {
    release(state, x.id, t)
    sp.wait.set(x.id, { s: x.s, r1: x.r1, c1: x.c1, r2: x.r2, c2: x.c2 })
  }
  place(state, f, a, area, t)
  return shown(a.at(0, 0))
}

/** Writes the array over its range, in place of the formula's previous spill. */
function place(state: EngineState, f: Formula, a: Arr, area: Area, t?: Tracker) {
  const sp = state.spills
  const sh = state.ctx.sheets[f.s]
  const { h, w, fill } = listing(a)
  allot(h * w)
  const old = sp.on.get(f.id)
  note(state, f.s, f.r, f.c, t)
  // cells of the old spill that the new one does not store, then the new ones (read before the old fills go)
  if (old) eachCell(old, (r, c) => (r >= f.r + h || c >= f.c + w) && clear(state, f.s, r, c, t))
  for (let r = 0; r < h; r++) {
    spend(w)
    for (let c = 0; c < w; c++) {
      if (r === 0 && c === 0) continue
      track(state, t, f.s, f.r + r, f.c + c)
      put(sh, f.r + r, f.c + c, shown(a.at(r, c)))
    }
  }
  if (old) sh.fills = sh.fills.filter((x) => !old.fills.includes(x))
  const fills: Fill[] = []
  if (fill !== undefined) {
    if (h < a.rows) fills.push({ r1: f.r + h, c1: f.c, r2: area.r2, c2: area.c2, v: fill })
    if (w < a.cols) fills.push({ r1: f.r, c1: f.c + w, r2: f.r + h - 1, c2: area.c2, v: fill })
    sh.fills.push(...fills)
  }
  sp.on.set(f.id, { ...area, id: f.id, h, w, fills })
  if (old) invalidate(state.ctx, old)
  invalidate(state.ctx, area)
  claim(state, f.id, area)
  sp.wait.delete(f.id)
  sp.why.delete(f.id)
  if (old && !equalAreas(old, area)) sp.touched.push({ ...old, step: sp.step })
  sp.touched.push({ ...area, step: sp.step })
}

/** The formula's array does not spill: #SPILL!, with the reason and, unless the edge stops it, the range it waits for. */
function blocked(state: EngineState, f: Formula, t: Tracker | undefined, why: string, wait?: Area): CErr {
  endSpill(state, f, t, false)
  if (wait) state.spills.wait.set(f.id, wait)
  state.spills.why.set(f.id, why)
  return E.SPILL
}

/**
 * The formula has no array to spill now (one value, an error, a cycle, or it is gone). In a cycle its claim
 * stays, so the cycle it closes through its range stays found.
 */
export function endSpill(state: EngineState, f: Formula, t: Tracker | undefined, cycle: boolean) {
  release(state, f.id, t)
  if (!cycle) claim(state, f.id, undefined)
  state.spills.wait.delete(f.id)
  state.spills.why.delete(f.id)
}

/** Sets the area the graph puts the readers of after the formula. */
function claim(state: EngineState, id: number, x: Area | undefined) {
  const sp = state.spills
  if (equalAreas(sp.claims.get(id), x)) return
  if (x) sp.claims.set(id, x)
  else sp.claims.delete(id)
  sp.claimsChanged = true
}

/** Takes a spill back: its cells are blank again (but those that hold input now). */
export function release(state: EngineState, id: number, t?: Tracker) {
  const sp = state.spills
  const x = sp.on.get(id)
  if (!x) return
  note(state, x.s, x.r1, x.c1, t)
  eachCell(x, (r, c) => clear(state, x.s, r, c, t))
  const sh = state.ctx.sheets[x.s]
  sh.fills = sh.fills.filter((f) => !x.fills.includes(f))
  sp.on.delete(id)
  invalidate(state.ctx, x)
  sp.touched.push({ ...x, step: sp.step })
}

/** Each stored cell of a spill but its formula's own. */
function eachCell(x: Spill, fn: (r: number, c: number) => unknown) {
  for (let r = x.r1; r < x.r1 + x.h; r++) for (let c = x.c1; c < x.c1 + x.w; c++) if (r !== x.r1 || c !== x.c1) fn(r, c)
}

function clear(state: EngineState, s: number, r: number, c: number, t?: Tracker) {
  if (state.spills.occupied.has(s, r, c)) return
  track(state, t, s, r, c)
  put(state.ctx.sheets[s], r, c, null)
}

/** Records the spill of a formula's cell before it first changes. */
function note(state: EngineState, s: number, r: number, c: number, t?: Tracker) {
  const k = key(s, r, c)
  if (t && !t.spills.has(k)) t.spills.set(k, describeAt(state, k))
}

/** A formula's spill in short (`3×1`, with the value of the unstored cells), '' when none: what changes compare. */
export function describeAt(state: EngineState, k: string): string {
  const id = state.at.get(k)
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (!x) return ''
  const fill = x.fills[0]?.v
  return `${x.r2 - x.r1 + 1}×${x.c2 - x.c1 + 1}${fill === undefined ? '' : ` ${canonValue(toCellValue(fill))}`}`
}

/** `E2#`: the range E2's array spilled into, E2 alone when it holds a formula that did not spill, else #REF!. */
export function spillRef(state: EngineState, s: number, r: number, c: number): Rng | CErr {
  const k = key(s, r, c)
  const id = state.at.get(k)
  if (id === undefined) return state.invalid.has(k) ? new Rng(s, r, c, r, c) : E.REF
  const x = state.spills.on.get(id)
  return x ? new Rng(s, x.r1, x.c1, x.r2, x.c2) : new Rng(s, r, c, r, c)
}

/** The spill a cell belongs to, but as its formula's own cell. */
export function ownerOf(state: EngineState, s: number, r: number, c: number): Spill | undefined {
  for (const x of state.spills.on.values()) if (x.s === s && r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2 && (r !== x.r1 || c !== x.c1)) return x
  return undefined
}
