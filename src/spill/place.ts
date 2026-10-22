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
import { allotSpill, spend } from '../exec/budget.js'
import { invalidate, put, same, shown, track, type Tracker } from '../evaluate.js'
import { parseMerge } from '../layout/merges.js'
import { formatAddress, MAX_COLS, MAX_ROWS, placeOf } from '../model.js'
import { hasCell, meets, type Rect } from '../structure/area.js'
import { quoteSheet } from '../syntax.js'
import { tableAt } from '../tables/model.js'
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
  for (const o of rivals(state, f.id, area)) giveWay(state, o, t)
  place(state, f, a, area, t)
  return shown(a.at(0, 0))
}

/** The array `a` of formula `f` spills if it can; the formula's own value (its first element, or #SPILL!). */
export function spill(state: EngineState, f: Formula, a: Arr, t?: Tracker): V {
  const sp = state.spills
  const area: Area = { s: f.s, r1: f.r, c1: f.c, r2: f.r + a.rows - 1, c2: f.c + a.cols - 1 }
  const cannot = `the array of ${a.rows}×${a.cols} values cannot spill`
  if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS) return blocked(state, f, t, `${cannot}: it would go past the edge of the sheet`)
  // Excel spills no array inside a table, nor into one
  const inside = tableAt(state.tables, f.s, f.r, f.c) ?? state.tables.list.find((x) => x.s === f.s && meets(x, area))
  if (inside) return blocked(state, f, t, `${cannot}: ${meets(inside, { r1: f.r, c1: f.c, r2: f.r, c2: f.c }) ? 'its cell is in' : 'it overlaps'} the table ${inside.name}`)
  const merge = mergesOf(state, f.s).find((m) => meets(m, area))
  if (merge) return blocked(state, f, t, `${cannot}: it overlaps the merged area ${cellText(state, f.s, merge.r1, merge.c1)}:${formatAddress(merge.r2, merge.c2)}`, area)
  const cell = sp.occupied.first(f.s, area, f.r, f.c)
  if (cell) return blocked(state, f, t, `${cannot}: ${cellText(state, f.s, cell[0], cell[1])} is not empty`, area)
  const over = rivals(state, f.id, area)
  // an old-kind array formula's area is never spilled over
  const first = over.find((x) => state.formulas[x.id]?.area || !beats(state, f.id, x.id, area))
  if (first) {
    const what = sp.on.has(first.id) ? 'the spilled array' : 'the range kept in a circular reference by the array'
    return blocked(state, f, t, `${cannot}: it overlaps ${what} of ${cellText(state, f.s, first.r1, first.c1)}`, area, first.id)
  }
  // the arrays this one wins over stop spilling and wait for their range
  for (const x of over) giveWay(state, x, t, f.id)
  place(state, f, a, area, t)
  return shown(a.at(0, 0))
}

/**
 * The arrays over `area` but formula `id`'s: those spilled there, and the ranges that formulas a circular reference
 * stops keep (their claims, `endSpill`), which an array spills over only where it wins (`beats`), as over a spill.
 */
function rivals(state: EngineState, id: number, area: Area): (Area & { id: number })[] {
  const sp = state.spills
  const out: (Area & { id: number })[] = [...sp.on.values()].filter((x) => x.id !== id && x.s === area.s && meets(x, area))
  for (const [k, x] of sp.claims) if (k !== id && x.s === area.s && !sp.on.has(k) && !sp.wait.has(k) && meets(x, area)) out.push({ ...x, id: k })
  return out
}

/** Array `x.id` gives way to formula `by` (or to an old-kind array's area): its spill, or the range a cycle keeps it, goes, and it waits for that range. */
function giveWay(state: EngineState, x: Area & { id: number }, t: Tracker | undefined, by?: number) {
  const sp = state.spills
  release(state, x.id, t)
  // the range it no longer holds closes no cycle: a claim left until the formula is computed again would stop it
  claim(state, x.id, undefined)
  sp.wait.set(x.id, { s: x.s, r1: x.r1, c1: x.c1, r2: x.r2, c2: x.c2 })
  if (by !== undefined) sp.by.set(x.id, by)
}

/** Writes the array over its range, in place of the formula's previous spill. */
function place(state: EngineState, f: Formula, a: Arr, area: Area, t?: Tracker) {
  const sp = state.spills
  const sh = state.ctx.sheets[f.s]
  const { h, w, fill } = listing(a)
  allotSpill(h * w)
  const old = sp.on.get(f.id)
  const fills: Fill[] = []
  if (fill !== undefined) {
    if (h < a.rows) fills.push({ r1: f.r + h, c1: f.c, r2: area.r2, c2: area.c2, v: fill })
    if (w < a.cols) fills.push({ r1: f.r, c1: f.c + w, r2: f.r + h - 1, c2: area.c2, v: fill })
  }
  // the same array placed again over the same cells changes nothing: its readers and the arrays waiting are not concerned
  let changed = !old || !equalAreas(old, area) || old.h !== h || old.w !== w || !sameFills(old.fills, fills)
  note(state, f.s, f.r, f.c, t)
  // cells of the old spill that the new one does not store, then the new ones (read before the old fills go)
  if (old) eachCell(old, (r, c) => (r >= f.r + h || c >= f.c + w) && clear(state, f.s, r, c, t))
  for (let r = 0; r < h; r++) {
    spend(w)
    for (let c = 0; c < w; c++) {
      if (r === 0 && c === 0) continue
      const v = shown(a.at(r, c))
      if (!changed && !same(sh.cols[f.c + c]?.get(f.r + r) ?? null, v)) changed = true
      track(state, t, f.s, f.r + r, f.c + c)
      put(sh, f.r + r, f.c + c, v)
    }
  }
  if (old) sh.fills = sh.fills.filter((x) => !old.fills.includes(x))
  sh.fills.push(...fills)
  sp.on.set(f.id, { ...area, id: f.id, h, w, fills })
  if (old) invalidate(state.ctx, old)
  invalidate(state.ctx, area)
  claim(state, f.id, area)
  sp.wait.delete(f.id)
  sp.why.delete(f.id)
  sp.by.delete(f.id)
  if (old && !equalAreas(old, area)) sp.touched.push({ ...old, step: sp.step })
  // an array that gave way to this one and now ranks before it (the graph changed) takes the cells: it is computed again
  // (an old-kind array formula gives way to none)
  if (!changed && !f.area) for (const [id, by] of sp.by) if (by === f.id && beats(state, id, f.id)) changed = true
  if (changed) sp.touched.push({ ...area, step: sp.step })
}

/** Two spills' fills: the same rectangles holding the same values. */
const sameFills = (a: Fill[], b: Fill[]) => a.length === b.length && a.every((x, i) => equalAreas({ s: 0, ...x }, { s: 0, ...b[i] }) && same(x.v, b[i].v))

/** The formula's array does not spill: #SPILL!, with the reason, unless the edge stops it the range it waits for, and the spill it gives way to. */
function blocked(state: EngineState, f: Formula, t: Tracker | undefined, why: string, wait?: Area, by?: number): CErr {
  endSpill(state, f, t, false)
  if (wait) state.spills.wait.set(f.id, wait)
  if (by !== undefined) state.spills.by.set(f.id, by)
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
  state.spills.by.delete(f.id)
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
  const k = placeOf(s, r, c)
  if (t && !t.spills.has(k)) t.spills.set(k, describeAt(state, k))
}

/** A formula's spill in short (`3×1`, with the value of the unstored cells), '' when none: what changes compare. */
export function describeAt(state: EngineState, k: number): string {
  const id = state.at.get(k)
  const x = id === undefined ? undefined : state.spills.on.get(id)
  if (!x) return ''
  const fill = x.fills[0]?.v
  return `${x.r2 - x.r1 + 1}×${x.c2 - x.c1 + 1}${fill === undefined ? '' : ` ${canonValue(toCellValue(fill))}`}`
}

/** `E2#`: the range E2's array spilled into, E2 alone when it holds a formula that did not spill, else #REF!. */
export function spillRef(state: EngineState, s: number, r: number, c: number): Rng | CErr {
  const id = state.at.get(placeOf(s, r, c))
  if (id === undefined) return state.invalid.has(key(s, r, c)) ? new Rng(s, r, c, r, c) : E.REF
  const x = state.spills.on.get(id)
  return x ? new Rng(s, x.r1, x.c1, x.r2, x.c2) : new Rng(s, r, c, r, c)
}

/** The spill a cell belongs to, but as its formula's own cell. */
export function ownerOf(state: EngineState, s: number, r: number, c: number): Spill | undefined {
  for (const x of state.spills.on.values()) if (x.s === s && hasCell(x, r, c) && (r !== x.r1 || c !== x.c1)) return x
  return undefined
}
