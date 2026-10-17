/**
 * Formulas a circular reference stops (`CellInput.held`, `docs/cycles.md`): without iterative calculation a formula in a
 * cycle, or one that reads a cycle, is not computed, as Excel computes none of them, and holds a value: the one its input
 * gives; else, in a book computed afresh, the one its file stores (`cached`) or 0; 0 for a formula the edits being
 * applied entered; else what it showed when the cycle came into its way. A formula out of a cycle's way is computed and
 * holds nothing. The state keeps the held values by formula (`EngineState.held`) and writes each change into the
 * formula's input, so a session's book gives a full computation what the session shows; a book put back (an undo) holds
 * what its inputs give, as a full computation of it does.
 */
import type { CachedValues, CellReplace, CellValue } from './contract.js'
import { toCellValue, type EngineState, type Formula } from './engine.js'
import { putInput } from './looks/sheet.js'
import { formatAddress, parseAddress, placeOf } from './model.js'
import { inverse, placer, type Geometry, type Place } from './structure/geometry.js'
import { sheetKey } from './syntax.js'
import { E, type V } from './values.js'

/** Whether formula `id` is one a circular reference stops: in a cycle or over one, in a book without iterative calculation. */
export const stopped = (state: EngineState, id: number): boolean => state.topo[id] < 0 && !state.input.iteration

/** A held value as the engine holds it. */
export const heldValue = (v: Exclude<CellValue, null>): V => (typeof v === 'object' ? E[v.error] : v)

/**
 * The cell's input with `held` set to `v` or left out (undefined); the first change of a cell's input goes to the
 * caller's journal (`EngineState.journal`). While the state is built, a sheet's cells are copied before their first
 * change (`EngineState.cow`), as the book it was given may share them.
 */
export function putHeld(state: EngineState, s: number, addr: string, v: V | undefined) {
  let sheet = state.input.sheets[s]
  const was = sheet.cells[addr]
  if (!was || (v === undefined && was.held === undefined)) return
  if (state.cow && !state.cow.has(s)) {
    state.cow.add(s)
    sheet = state.input.sheets[s] = { ...sheet, cells: { ...sheet.cells } }
  }
  const { held: _, ...rest } = was
  const { row, col } = parseAddress(addr)
  putInput(sheet, addr, row, col, v === undefined ? rest : { ...rest, held: toCellValue(v) as Exclude<CellValue, null> })
  const k = `${s}!${addr}`
  if (state.journal && !state.journal.has(k)) state.journal.set(k, { sheet: state.sheetNames[s], cell: addr, input: was })
}

/** Formula `f`'s held value set (`v`), or dropped (undefined) as it comes out of a cycle's way. */
export function setHeld(state: EngineState, f: Formula, v: V | undefined) {
  if (v === undefined) state.held.delete(f.id)
  else state.held.set(f.id, v)
  putHeld(state, f.s, formatAddress(f.r, f.c), v)
}

/**
 * The value of formula `f`, which a circular reference stops, `shows` what its cell shows now: its held value; else what
 * it takes as the cycle comes into its way (`EngineState.anew`), which it holds from now on where that is not 0.
 */
export function heldAt(state: EngineState, f: Formula, shows: V): V {
  const v = state.held.get(f.id)
  if (v !== undefined) return v
  const a = state.anew
  const w = a === undefined ? shows : a instanceof Set ? (a.has(f.id) ? 0 : shows) : (a(f) ?? 0)
  if (w !== 0) setHeld(state, f, w)
  return w
}

/** The formulas computed this time that hold a value, out of a cycle's way once the computation settles, drop it. */
export function unhold(state: EngineState) {
  for (const id of state.outside) if (state.formulas[id] && state.held.has(id) && !stopped(state, id)) setHeld(state, state.formulas[id]!, undefined)
  state.outside.clear()
}

/**
 * What a formula a circular reference stops takes, holding nothing, in a book built (`EngineState.anew`): what `prior`
 * gives (what it showed before a session's book was built again), else the value `cached` stores for it, empty text
 * where it stores none; else 0.
 */
export const builtAnew = (state: EngineState, cached?: CachedValues, prior?: (f: Formula) => V | undefined): ((f: Formula) => V | undefined) =>
  prior ?? (cached ? storedIn(state, cached) : () => undefined)

/**
 * What the formulas of a book built again from `old`'s showed (`EngineOptions.prior`): each one's value in the cell the
 * geometry `g` brought it from, where that held a formula no cycle stopped; one a cycle stopped holds its own already.
 */
export function priorOf(old: EngineState, g: Geometry): (f: Formula) => V | undefined {
  const back = placer(inverse(g))
  const p: Place = { s: 0, r: 0, c: 0 }
  return (f) => {
    p.s = f.s
    p.r = f.r
    p.c = f.c
    if (!back(p)) return undefined
    const id = old.at.get(placeOf(p.s, p.r, p.c))
    return id === undefined || stopped(old, id) ? undefined : (old.ctx.sheets[p.s].cols[p.c]?.[p.r] ?? undefined)
  }
}

/** The value a file stores for a formula, by the sheets of `cached` that the book has: empty text where it stores none. */
function storedIn(state: EngineState, cached: CachedValues): (f: Formula) => V {
  const sheets = new Map(cached.sheets.map((x) => [state.sheetIdx.get(sheetKey(x.name)), x.cells]))
  return (f) => {
    const v = sheets.get(f.s)?.[formatAddress(f.r, f.c)]?.value
    return v === undefined || v === null ? '' : heldValue(v)
  }
}

/** The inputs `was` puts back, and those the engine changed besides (`EngineState.journal`) of the cells `was` does not hold. */
export function journaled(was: CellReplace[], journal: Map<string, CellReplace> | undefined): CellReplace[] {
  if (!journal?.size) return was
  const seen = new Set(was.map((e) => `${sheetKey(e.sheet)}!${e.cell}`))
  return [...was, ...[...journal.values()].filter((e) => !seen.has(`${sheetKey(e.sheet)}!${e.cell}`))]
}
