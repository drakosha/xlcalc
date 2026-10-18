/**
 * The formulas of conditional formats computed at a cell of the rule's ranges. A formula is written for the top-left
 * cell of the rule's first range and moves from there to each cell as a copied formula does (`origin` of the walk).
 * Where only its references and the functions of their cell move it, one code serves every cell and reads the cell it
 * runs at (`here` of the walk); otherwise its code is made once per cell along the axes it moves and kept with the
 * session's state up to `KEEP` bytes, the code read least recently going first.
 */
import { codegenContext, type EngineState } from '../engine.js'
import { closureBackend } from '../exec/closures.js'
import { walk } from '../exec/walk.js'
import { nameFor } from '../names.js'
import { MAX_COLS } from '../model.js'
import { decode } from '../structure/area.js'
import { FUNCTIONS } from '../functions.js'
import { sheetKey, validate, walk as walkTree, type Ast } from '../syntax.js'
import { tableNames } from '../tables/model.js'
import { Areas, Rng, type V } from '../values.js'
import { Arr } from '../arrays/value.js'
import type { Arg } from '../args.js'

/** A formula checked against the book as a cell's on sheet `s` is: its tree, or null when it does not parse or names what the book lacks. */
export function prepare(state: EngineState, s: number, text: string): Ast | null {
  const known = (name: string, sheet: string | undefined) => !!nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)
  const r = validate(text, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables: tableNames(state.tables) })
  return r.ok ? r.ast : null
}

/** A formula's code at cell (r, c) of sheet `s`, written for `origin`: closures, whatever the book's execution mode; with `here`, the code of every cell, which reads the cell it runs at there. */
function compileAt(state: EngineState, s: number, ast: Ast, origin: { r: number; c: number }, r: number, c: number, here?: { r: number; c: number }): () => Arg {
  return walk({ ...codegenContext(state, s), origin, here }, ast, s, r, c, closureBackend(state.rt)).code
}

/** The axes along which a reference has a bound without `$`. */
function axes(ref: Extract<Ast, { k: 'ref' | 'ref3d' | 'xref' }>['ref']): [boolean, boolean] {
  const a = decode(ref)
  return [!a.r.lo.abs || !a.r.hi.abs, !a.c.lo.abs || !a.c.hi.abs]
}

/** A reference with a bound without `$`. */
const moving = (n: Ast) => (n.k === 'ref' || n.k === 'ref3d' || n.k === 'xref') && axes(n.ref).some(Boolean)

/**
 * How a formula's code differs from cell to cell: the axes along which it moves (relative references, names, `@`, ROW()
 * and the other functions of their cell), and whether one code that reads the cell it runs at serves every cell
 * (`shared`): not with a name, `@`, a structured reference, a relative reference across sheets or books, or LOOKUP's
 * result widened by a moving reference, which the code of a cell decides there; and its nodes.
 */
function movesAlong(ast: Ast): { rows: boolean; cols: boolean; shared: boolean; nodes: number } {
  const out = { rows: false, cols: false, shared: true, nodes: 0 }
  walkTree(ast, (n) => {
    out.nodes++
    if (n.k === 'ref' || n.k === 'ref3d') {
      const [rows, cols] = axes(n.ref)
      out.rows ||= rows
      out.cols ||= cols
    } else if (n.k === 'name' || n.k === 'at' || n.k === 'table' || (n.k === 'fn' && FUNCTIONS[n.name]?.cell)) out.rows = out.cols = true
    if (n.k === 'name' || n.k === 'at' || n.k === 'table' || (n.k !== 'ref' && moving(n))) out.shared = false
    else if (n.k === 'fn' && n.name === 'LOOKUP' && n.args.length === 3 && n.args.slice(1).some(moving)) out.shared = false
  })
  return out
}

/** A cell's code kept: in its formula's map by its cell, with its bytes, and between the codes read before and after it. */
interface Code {
  fn: () => Arg
  fns: Map<number, Code>
  i: number
  bytes: number
  prev?: Code
  next?: Code
}

/** Per state: the formulas as functions of the cell (`evaluator`), by sheet, origin and text; the code per cell kept. */
interface Kept {
  formulas: Map<string, (r: number, c: number) => V>
  /** The ends of the list of the code per cell kept: the code read least recently and the code read last. */
  oldest?: Code
  newest?: Code
  /** Its bytes. */
  held: number
}
const kept = new WeakMap<EngineState, Kept>()

/** The code per cell a state keeps, in bytes: past them the code read least recently goes. */
const KEEP = 8 * 2 ** 20

/** Bytes of a cell's code besides its nodes', and per node of its formula, as measured on Node (`A1>lim` 1.1 KB, 39 nodes 5.9 KB). */
const CODE_BYTES = 720
const NODE_BYTES = 135

/**
 * A formula of sheet `s` written for `origin`, as a function of the cell: its value there. Its code is made once for all
 * cells where only its references and the functions of their cell move it, or nothing does; otherwise once per cell
 * along the axes it moves (`movesAlong`), kept with the state up to `KEEP` bytes, the code read least recently going
 * first. Null when the formula does not parse or names what the book lacks.
 */
export function evaluator(state: EngineState, s: number, ast: Ast, text: string, origin: { r: number; c: number }): (r: number, c: number) => V {
  let m = kept.get(state)
  if (!m) kept.set(state, (m = { formulas: new Map(), held: 0 }))
  const k = `${s}\n${origin.r}:${origin.c}\n${text}`
  let f = m.formulas.get(k)
  if (!f) m.formulas.set(k, (f = made(state, s, ast, origin, m)))
  return f
}

/** The formula as a function of the cell (`evaluator`), its code made as `movesAlong` says. */
function made(state: EngineState, s: number, ast: Ast, origin: { r: number; c: number }, m: Kept): (r: number, c: number) => V {
  const { rows, cols, shared, nodes } = movesAlong(ast)
  if (!rows && !cols) {
    const fn = compileAt(state, s, ast, origin, origin.r, origin.c)
    return () => first(state, fn())
  }
  if (shared) {
    const here = { r: origin.r, c: origin.c }
    const fn = compileAt(state, s, ast, origin, origin.r, origin.c, here)
    return (r, c) => {
      here.r = r
      here.c = c
      return first(state, fn())
    }
  }
  const fns = new Map<number, Code>()
  const bytes = CODE_BYTES + NODE_BYTES * nodes
  return (r, c) => {
    const i = (rows ? r : 0) * MAX_COLS + (cols ? c : 0)
    let code = fns.get(i)
    if (!code) {
      fns.set(i, (code = { fn: compileAt(state, s, ast, origin, rows ? r : origin.r, cols ? c : origin.c), fns, i, bytes }))
      m.held += bytes
    }
    if (code !== m.newest) read(m, code)
    if (m.held > KEEP) drop(m)
    return first(state, code.fn())
  }
}

/** A cell's code just read goes to the end of the state's list, from its place in it or as new. */
function read(m: Kept, code: Code) {
  if (code.prev) code.prev.next = code.next
  else if (m.oldest === code) m.oldest = code.next
  if (code.next) code.next.prev = code.prev
  code.prev = m.newest
  code.next = undefined
  if (m.newest) m.newest.next = code
  else m.oldest = code
  m.newest = code
}

/** The code read least recently goes until the state keeps `KEEP` bytes; the code read last stays. */
function drop(m: Kept) {
  while (m.held > KEEP && m.oldest !== m.newest) {
    const code = m.oldest!
    m.oldest = code.next
    m.oldest!.prev = undefined
    code.fns.delete(code.i)
    m.held -= code.bytes
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
