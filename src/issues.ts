/**
 * Error reasons: why each cell with an error has it (a formula check, a cycle, an array that cannot spill, an
 * error from another cell or its own, an error element of a spilled array), and the cell an error began in.
 */
import type { CellIssueDetail, CellRef, CheckCode, ErrorCode } from './contract.js'
import { key, valueAt, type EngineState, type Formula } from './engine.js'
import { formatAddress } from './model.js'
import { ownerOf } from './spill/place.js'
import type { Spill } from './spill/state.js'
import { quoteSheet } from './syntax.js'
import { isErr } from './values.js'

const GENERIC: Record<ErrorCode, string> = {
  DIV0: 'division by zero',
  NA: 'value not found',
  NAME: 'unknown function or name',
  NUM: 'number out of range',
  REF: 'invalid reference',
  VALUE: 'wrong type of value',
  NULL: 'ranges do not intersect',
  CYCLE: 'circular reference',
  SPILL: 'array result does not fit',
  CALC: 'empty array or text too long',
  ERROR: 'error',
}

export interface Issue {
  s: number
  r: number
  c: number
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
}

/** A cell of a sheet: zero-based row and column. */
export interface Place {
  s: number
  r: number
  c: number
}

/** A cell's reference text: `'Sheet'!B3`. */
export const placeText = (state: EngineState, p: Place) => `${quoteSheet(state.sheetNames[p.s])}!${formatAddress(p.r, p.c)}`

/** The issue of a formula whose value is the error `code`. */
function formulaIssue(state: EngineState, f: Formula, code: ErrorCode): Issue {
  const why = state.spills.why.get(f.id)
  if (code === 'SPILL' && why) return { s: f.s, r: f.r, c: f.c, error: code, reason: why }
  let reason = GENERIC[code]
  if (code !== 'CYCLE') {
    const from = errorSource(state, f, code)
    if (from) reason = `${GENERIC[code]}, from ${placeText(state, from)}`
    else if (code === 'NA' && f.external) reason = 'value not found (or no value for an external call)'
  }
  return { s: f.s, r: f.r, c: f.c, error: code, reason }
}

/** The reason of a spilled cell (not its formula's) holding the error `code`, by its formula's reference text. */
export const spilledReason = (code: ErrorCode, from: string) => `${GENERIC[code]}, in the array spilled from ${from}`

/** The issue of a spilled cell (not its formula's) holding the error `code`. */
const spilledIssue = (state: EngineState, x: Spill, r: number, c: number, code: ErrorCode): Issue => ({ s: x.s, r, c, error: code, reason: spilledReason(code, placeText(state, { s: x.s, r: x.r1, c: x.c1 })) })

/** The cell a reason names as the one its error came from (`, from 'S'!A2`, `in the array spilled from 'S'!B1`), or null. */
export function reasonSource(reason: string): CellRef | null {
  const m = /, (?:in the array spilled )?from '((?:[^']|'')+)'!([A-Z]+[1-9]\d*)$/.exec(reason)
  return m ? { sheet: m[1].replaceAll("''", "'"), cell: m[2] } : null
}

/** Cells with an error and the reason. */
export function issues(state: EngineState): Issue[] {
  const out: Issue[] = []
  for (const [k, inv] of state.invalid) {
    const [s, r, c] = k.split(':').map(Number)
    out.push({ s, r, c, error: inv.error, reason: inv.message, check: inv.code, at: inv.at })
  }
  for (const f of state.formulas) {
    if (!f) continue
    const v = valueAt(state, f.s, f.r, f.c)
    if (isErr(v)) out.push(formulaIssue(state, f, v.code))
  }
  // the stored cells of spilled arrays (the unstored rest of a large one is its formula's)
  for (const x of state.spills.on.values())
    for (let r = x.r1; r < x.r1 + x.h; r++)
      for (let c = x.c1; c < x.c1 + x.w; c++) {
        const v = valueAt(state, x.s, r, c)
        if (isErr(v) && (r !== x.r1 || c !== x.c1)) out.push(spilledIssue(state, x, r, c, v.code))
      }
  return out.sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
}

/** The issue of one cell, as `issues` lists it (a cell of a large spill's unstored rest too); null for a cell without an error. */
export function cellIssue(state: EngineState, s: number, r: number, c: number): Issue | null {
  const inv = state.invalid.get(key(s, r, c))
  if (inv) return { s, r, c, error: inv.error, reason: inv.message, check: inv.code, at: inv.at }
  const v = valueAt(state, s, r, c)
  if (!isErr(v)) return null
  const id = state.at.get(key(s, r, c))
  if (id !== undefined) return formulaIssue(state, state.formulas[id]!, v.code)
  const x = ownerOf(state, s, r, c)
  return x ? spilledIssue(state, x, r, c, v.code) : null
}

/** The first cell among a formula's references holding the error `code`, the one its reason names, with the index of its reference. */
function errorSource(state: EngineState, f: Formula, code: ErrorCode): (Place & { ref: number }) | null {
  for (let k = 0; k < f.refs.length; k++) {
    const g = f.refs[k]
    const flat = state.rt.flat(g)
    const b = state.rt.bounds(g)
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i]
      if (isErr(v) && v.code === code) return { s: g.s, r: b.r1 + Math.floor(i / b.w), c: b.c1 + (i % b.w), ref: k }
    }
  }
  return null
}

/**
 * The next cell on the way back from a cell holding the error `code` to the cell it began in: for a formula, the cell its
 * reason names (with the index of the reference it is in); for a spilled cell, its array's formula; null where it began.
 */
export function errorStep(state: EngineState, p: Place, code: ErrorCode): (Place & { ref?: number }) | null {
  if (state.invalid.has(key(p.s, p.r, p.c))) return null
  const id = state.at.get(key(p.s, p.r, p.c))
  if (id === undefined) {
    const x = ownerOf(state, p.s, p.r, p.c)
    return x ? { s: x.s, r: x.r1, c: x.c1 } : null
  }
  const f = state.formulas[id]!
  return code === 'SPILL' && state.spills.why.has(f.id) ? null : errorSource(state, f, code)
}

/**
 * The cell an error of `code` at `p` began in: the last of its steps, or in a cycle the first cell the way comes back to.
 * `memo` keeps the origins of ways without a cycle, which every cell on such a way shares, for the next calls.
 */
export function errorOrigin(state: EngineState, p: Place, code: ErrorCode, memo?: Map<string, Place>): Place {
  return wayBack(p, code, (at) => errorStep(state, at, code), memo)
}

/** The last of the steps from `p` (`step`, null where the error began), or in a cycle the first cell the way comes back to; `memo` as `errorOrigin`'s. */
export function wayBack(p: Place, code: ErrorCode, step: (at: Place) => Place | null, memo?: Map<string, Place>): Place {
  const way = new Set<string>()
  let at = p
  for (;;) {
    const k = `${code}:${key(at.s, at.r, at.c)}`
    const known = memo?.get(k)
    if (known) return done(known)
    way.add(k)
    const next = step(at)
    if (!next) return done(at)
    if (way.has(`${code}:${key(next.s, next.r, next.c)}`)) return { s: next.s, r: next.r, c: next.c }
    at = next
  }
  function done(origin: Place): Place {
    for (const k of way) memo?.set(k, origin)
    return origin
  }
}

/** An issue as `Session.getCellIssue` gives it, with the cell its error began in. */
export const issueWith = (i: Omit<Issue, 's' | 'r' | 'c'>, origin: CellRef): CellIssueDetail => ({ error: i.error, reason: i.reason, ...(i.check ? { check: i.check, at: i.at } : {}), origin })

/** A cell's issue with the cell its error began in (`Session.getCellIssue`); null for a cell without an error. `memo` as `errorOrigin`'s. */
export function issueDetail(state: EngineState, s: number, r: number, c: number, memo?: Map<string, Place>): CellIssueDetail | null {
  const i = cellIssue(state, s, r, c)
  if (!i) return null
  const o = errorOrigin(state, { s, r, c }, i.error, memo)
  return issueWith(i, { sheet: state.sheetNames[o.s], cell: formatAddress(o.r, o.c) })
}
