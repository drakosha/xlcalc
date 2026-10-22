/** What the engine keeps about spilled arrays (`docs/spill.md`): where they spilled, what waits, why not. */
import type { Rect } from '../structure/area.js'
import type { Fill } from './fill.js'
import { Occupied } from './occupied.js'
import type { Places } from './rank.js'

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
  /** The formulas whose claim a cycle keeps without a spill (`endSpill`, an iteration): the ranges other arrays meet as a spill's. */
  kept: Set<number>
  /** Formulas whose array could not spill, with the area it needs: they are computed again when it changes. */
  wait: Map<number, Area>
  /** Why a formula's array could not spill. */
  why: Map<number, string>
  /** The spilled array that a waiting one gave way to, by the waiting one (`beats` may turn once the graph changes). */
  by: Map<number, number>
  occupied: Occupied
  /** Merged areas by sheet, parsed when first needed. */
  merges: Map<number, Rect[]>
  /** Of two arrays over one cell, the one with the lower rank spills (made when needed, dropped with the graph). */
  rank?: Map<number, number>
  /** Formulas that read a cell their array may spill over (`rank.ts`), as found, and the formulas by place; dropped with the graph. */
  looped?: Map<number, boolean>
  places?: Places
  /** Steps so far, counted on across rounds: a formula's evaluation, a change of spilled cells. */
  step: number
  /** The step the current round began after: a formula's step in `done` before it is of an earlier round. */
  since: number
  /** The step each formula (by id) was last computed at; a typed array, written once per formula computed, without a map's hashing. */
  done: Float64Array
  touched: (Area & { step: number })[]
  claimsChanged: boolean
  /** Iterative calculation is computing the cycles: their formulas keep their claims (`endSpill`). */
  iterating: boolean
}

export const newSpills = (): Spills => ({ on: new Map(), claims: new Map(), kept: new Set(), wait: new Map(), why: new Map(), by: new Map(), occupied: new Occupied(), merges: new Map(), step: 0, since: 0, done: new Float64Array(0), touched: [], claimsChanged: false, iterating: false })

export const equalAreas = (a: Area | undefined, b: Area | undefined) => a === b || (!!a && !!b && a.s === b.s && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2)
