/** What the engine keeps about spilled arrays (`docs/spill.md`): where they spilled, what waits, why not. */
import type { Rect } from '../structure/area.js'
import type { Fill } from './fill.js'
import { Occupied } from './occupied.js'

/** A rectangle of a sheet. */
export interface Area extends Rect {
  s: number
}

/**
 * An array that spilled from its formula's cell, the area's top-left: `h` × `w` cells from there are stored
 * (listed), the rest of the area is `fills` holding one value (only for a large array).
 */
export interface Spill extends Area {
  id: number
  h: number
  w: number
  fills: Fill[]
}

export interface Spills {
  /** Formulas whose array spilled, by id. */
  on: Map<number, Spill>
  /** The areas whose readers the graph puts after their formula: a spill's, kept while its formula is in a cycle. */
  claims: Map<number, Area>
  /** Formulas whose array could not spill, with the area it needs: they are computed again when it changes. */
  wait: Map<number, Area>
  /** Why a formula's array could not spill. */
  why: Map<number, string>
  occupied: Occupied
  /** Merged areas by sheet, parsed when first needed. */
  merges: Map<number, Rect[]>
  /** Of two arrays over one cell, the one with the lower rank spills (made when needed, dropped with the graph). */
  rank?: Map<number, number>
  /** Steps so far, counted on across rounds: a formula's evaluation, a change of spilled cells. */
  step: number
  /** The step the current round began after: a formula's step in `done` before it is of an earlier round. */
  since: number
  /** The step each formula (by id) was last computed at; a typed array, written once per formula computed, without a map's hashing. */
  done: Float64Array
  touched: (Area & { step: number })[]
  claimsChanged: boolean
}

export const newSpills = (): Spills => ({ on: new Map(), claims: new Map(), wait: new Map(), why: new Map(), occupied: new Occupied(), merges: new Map(), step: 0, since: 0, done: new Float64Array(0), touched: [], claimsChanged: false })

export const equalAreas = (a: Area | undefined, b: Area | undefined) => a === b || (!!a && !!b && a.s === b.s && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2)
