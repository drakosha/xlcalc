/**
 * The formulas of conditional formats computed at a cell of the rule's ranges. A formula is written for the top-left
 * cell of the rule's first range and moves from there to each cell as a copied formula does (`origin` of the walk);
 * its code is made once per cell along the axes it moves and kept with the session's state.
 */
import { codegenContext, type EngineState } from '../engine.js'
import { closureBackend } from '../exec/closures.js'
import { walk } from '../exec/walk.js'
import { nameFor } from '../names.js'
import { MAX_COLS } from '../model.js'
import { decode } from '../structure/area.js'
import { FUNCTIONS } from '../functions.js'
import { sheetKey, validate, walk as walkTree, type Ast } from '../syntax.js'
import { Areas, Rng, type V } from '../values.js'
import { Arr } from '../arrays/value.js'
import type { Arg } from '../args.js'

/** A formula checked against the book as a cell's on sheet `s` is: its tree, or null when it does not parse or names what the book lacks. */
export function prepare(state: EngineState, s: number, text: string): Ast | null {
  const known = (name: string, sheet: string | undefined) => !!nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)
  const r = validate(text, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables: state.tables.list.map((t) => ({ name: t.name, columns: t.cols })) })
  return r.ok ? r.ast : null
}

/** A formula's code at cell (r, c) of sheet `s`, written for `origin`: closures, whatever the book's execution mode. */
function compileAt(state: EngineState, s: number, ast: Ast, origin: { r: number; c: number }, r: number, c: number): () => Arg {
  return walk({ ...codegenContext(state, s), origin }, ast, s, r, c, closureBackend(state.rt)).code
}

/** The axes along which a formula's code differs from cell to cell: relative references, names, `@`, ROW() and the other functions of their cell. */
function movesAlong(ast: Ast): { rows: boolean; cols: boolean } {
  const out = { rows: false, cols: false }
  walkTree(ast, (n) => {
    if (n.k === 'ref' || n.k === 'ref3d') {
      const a = decode(n.ref)
      out.rows ||= !a.r.lo.abs || !a.r.hi.abs
      out.cols ||= !a.c.lo.abs || !a.c.hi.abs
    } else if (n.k === 'name' || n.k === 'at' || n.k === 'table' || (n.k === 'fn' && FUNCTIONS[n.name]?.cell)) out.rows = out.cols = true
  })
  return out
}

/** Code kept per state, by sheet, formula and origin: each by the cell it is for, along the axes it moves (`movesAlong`). */
const kept = new WeakMap<EngineState, { held: number; formulas: Map<string, { rows: boolean; cols: boolean; fns: Map<number, () => Arg> }> }>()

/** At most this many pieces of code kept per state; past it a formula's are dropped and made again. */
const KEEP = 200_000

/**
 * A formula of sheet `s` written for `origin`, as a function of the cell: its value there. The code is made once per
 * cell along the axes it moves, once for all where it does not move, and kept with the state; null when the formula
 * does not parse or names what the book lacks.
 */
export function evaluator(state: EngineState, s: number, ast: Ast, text: string, origin: { r: number; c: number }): (r: number, c: number) => V {
  let m = kept.get(state)
  if (!m) kept.set(state, (m = { held: 0, formulas: new Map() }))
  const k = `${s}\n${origin.r}:${origin.c}\n${text}`
  let x = m.formulas.get(k)
  if (!x) m.formulas.set(k, (x = { ...movesAlong(ast), fns: new Map() }))
  const { rows, cols, fns } = x
  const all = m
  return (r, c) => {
    const i = (rows ? r : 0) * MAX_COLS + (cols ? c : 0)
    let fn = fns.get(i)
    if (!fn) {
      if (all.held >= KEEP) {
        all.held -= fns.size
        fns.clear()
      }
      fns.set(i, (fn = compileAt(state, s, ast, origin, rows ? r : origin.r, cols ? c : origin.c)))
      all.held++
    }
    return first(state, fn())
  }
}

/** One value of a formula's answer: a reference's or an array's top-left one. */
export function first(state: EngineState, a: Arg): V {
  if (a instanceof Rng) return state.rt.cell(a.s, a.r1, a.c1)
  if (a instanceof Areas) return first(state, a.list[0])
  if (a instanceof Arr) return a.v[0] ?? null
  return a as V
}

/**
 * Whether a formula's answer holds, read as IF reads its condition: TRUE, a number other than 0, the text of TRUE in the
 * book's locale (the set `cond/looks`: `IF($A1>3,"TRUE","")` holds); other text, a blank and an error do not.
 */
export const holds = (state: EngineState, v: V): boolean => v !== null && state.ctx.rules.toBool(v) === true
