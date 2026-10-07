/**
 * The one walk over a formula's AST that both execution modes share. References, names, shapes
 * (`arrays/shape.ts`), function specs, arguments left out and the formula's cell are resolved here, together
 * with the formula's references and flags; LET's names are the parser's slots, whose storage each backend keeps
 * per computation of the LET. A backend (`js.ts` for `eval`, `closures.ts`) only builds the code of each node
 * from what the walk decided, so the two modes cannot part on a decision.
 */
import { shapes, type NameRef } from '../arrays/shape.js'
import { Arr } from '../arrays/value.js'
import { FUNCTIONS } from '../functions.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { decode, type Area, type Bound, type Rect } from '../structure/area.js'
import { canonical, type Ast, type BinOp, type RangeRef } from '../syntax.js'
import { E, Rng, type V } from '../values.js'

/** A defined name as the walk expands it: its formula, which sees the global names (a sheet's own only as `'S'!name`). */
export interface NameDef {
  ast: Ast
}

/** What the walk needs to know of the book, seen from the formula's sheet. */
export interface CodegenContext {
  /** The index of a sheet (null: the formula's own). */
  sheet(name: string | null): number
  /** A name as the names of sheet `scope` show it: its local one there, else the global one (-1: the global one). */
  name(name: string, scope: number): NameDef | undefined
  /** The area of an array formula of the old kind: `ROW()` and `COLUMN()` give its rows and columns. */
  area?: Rect
}

export type Operator = 'add' | 'sub' | 'mul' | 'div' | 'pow' | 'cat'

/** The runtime's arithmetic and text operators by symbol; comparisons go through `cmp` (lifted: `C`). */
export const OPERATORS: Partial<Record<BinOp, Operator>> = { '+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat' }

/**
 * Code of one node in a mode. `lift`: an operand or a one-value argument can be an array, so the node goes
 * through the lifting operators (`L1`, `L2`) or callers (`FL`, `LZ`); otherwise it is plain scalar code.
 */
export interface Backend<T> {
  /** A literal: a number, text, a boolean, an error or a blank. */
  value(v: V): T
  /** An argument left out that the function takes as not given. */
  absent(): T
  /** An array constant, built once per compiled formula. */
  array(a: Arr): T
  cell(s: number, r: number, c: number): T
  /** A spill reference (`E2#`): the range the cell's array spilled into, the cell alone, or #REF!. */
  spill(s: number, r: number, c: number): T
  /** A constant range, built once per compiled formula; `one`: a single cell in a one-value position. */
  range(g: Rng, one: boolean): T
  unary(op: 'neg' | 'pct', a: T, lift: boolean): T
  binary(op: BinOp, a: T, b: T, lift: boolean): T
  call(name: string, args: T[], lift: boolean): T
  /** A lazy function (IF, IFERROR): each argument computed on demand; `check`, the positions that can lift it. */
  lazy(name: string, args: T[], check: number[]): T
  /** The value of a call that may answer a one-cell reference (CHOOSE, IF) where one value goes. */
  read(code: T): T
  /** A reference across sheets: its rectangle on each sheet, built once per compiled formula. */
  areas(list: Rng[]): T
  /** A union of references (#VALUE! unless they are references on one sheet). */
  union(items: T[]): T
  /** `@x` at row `r`, column `c`: a reference's cell there, an array's top-left value. */
  at(a: T, r: number, c: number): T
  /** An external function, by the canonical text of its call. */
  external(text: string): T
  /**
   * A LET of `count` names: `build` gets the code of a use of its i-th name and gives the code of the values
   * and of the result. A value is computed on its first use, once in each computation of the LET.
   */
  let(count: number, build: (local: (i: number) => T) => { values: T[]; body: T }): T
}

/** A LET value: computed on its first use, then kept. */
export function once<A>(f: () => A): () => A {
  let done = false
  let v: A
  return () => {
    if (!done) {
      v = f()
      done = true
    }
    return v
  }
}

/** What a formula reads and how it behaves, the same in both modes. */
export interface FormulaInfo {
  refs: Rng[]
  /** Calls a volatile function directly. */
  volatile: boolean
  /** Calls SUBTOTAL or AGGREGATE: those two skip the cell. */
  nested: boolean
  /** Calls an external function (its missing value is #N/A). */
  external: boolean
}

/** The rows (a column of their numbers) or the columns (a row) of an area. */
function lines(x: Rect, rows: boolean): Arr {
  const n = rows ? x.r2 - x.r1 + 1 : x.c2 - x.c1 + 1
  const v = Array.from({ length: n }, (_, i) => (rows ? x.r1 : x.c1) + i + 1)
  return rows ? new Arr(n, 1, n, 1, v) : new Arr(1, n, 1, n, v)
}

/** One axis of a name's relative bound used at index `at`: moved by it, wrapping round the sheet. */
const moved = (b: Bound, at: number, max: number) => (b.abs ? b.i : (b.i + at) % max)

/** The bounds of a name's references, read once: a name's formula is walked at every use. */
const bounds = new WeakMap<RangeRef, Area | null>()
function relative(ref: RangeRef): Area | null {
  let a = bounds.get(ref)
  if (a === undefined) {
    const d = decode(ref)
    a = d.r.lo.abs && d.r.hi.abs && d.c.lo.abs && d.c.hi.abs ? null : d
    bounds.set(ref, a)
  }
  return a
}

/** Walks a formula at sheet `s`, row `r`, column `c`: its information and the code of its root in `b`. */
export function walk<T>(ctx: CodegenContext, ast: Ast, s: number, r: number, c: number, b: Backend<T>): FormulaInfo & { code: T } {
  const info: FormulaInfo = { refs: [], volatile: false, nested: false, external: false }
  // a name as the names of `scope` show it, or as its sheet's (`'S'!rate`)
  const resolve: NameRef = (n, scope) => ctx.name(n.name, n.sheet === undefined ? scope : ctx.sheet(n.sheet))
  const shape = shapes(ast, s, resolve, !!ctx.area)
  const many = (n: Ast) => (shape.get(n) ?? 0) > 0
  // the code of a use of each LET name, by slot (a name's formula has its own LET names)
  let uses: T[] = []
  // whose names the walk sees, and the names it is inside (their references without `$` move by the cell from A1)
  let scope = s
  const inside = new Set<NameDef>()
  // a reference's rectangle: in a name's formula, its relative bounds moved by the formula's cell
  const rect = (n: Extract<Ast, { k: 'ref' | 'ref3d' }>) => {
    const a = inside.size ? relative(n.ref) : null
    if (!a) return n.ref
    const [r1, r2] = [moved(a.r.lo, r, MAX_ROWS), moved(a.r.hi, r, MAX_ROWS)]
    const [c1, c2] = [moved(a.c.lo, c, MAX_COLS), moved(a.c.hi, c, MAX_COLS)]
    return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) }
  }
  // an argument left out: a value, or not given
  const omitted = (v: V | undefined) => (v === undefined ? b.absent() : b.value(v))
  // a value that may be a reference, read as a value where one value goes
  const value = (code: T, n: Ast, rangeOk: boolean) => (!rangeOk && !many(n) ? b.read(code) : code)
  // rangeOk: a position where the function expects a range (a range object goes there, even for a single cell)
  const go = (n: Ast, rangeOk = false): T => {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
        return b.value(n.v)
      case 'err':
        return b.value(E[n.code])
      case 'arr': {
        const v = n.rows.flat().map((x) => (x.k === 'err' ? E[x.code] : x.v))
        return b.array(new Arr(n.rows.length, n.rows[0].length, n.rows.length, n.rows[0].length, v))
      }
      case 'ref': {
        const x = rect(n)
        const g = new Rng(ctx.sheet(n.sheet), x.r1, x.c1, x.r2, x.c2)
        info.refs.push(g)
        if (n.ref.spill) return b.spill(g.s, g.r1, g.c1)
        if (n.ref.kind === 'cell' && !rangeOk) return b.cell(g.s, g.r1, g.c1)
        return b.range(g, !rangeOk && g.r2 === g.r1 && g.c2 === g.c1)
      }
      case 'ref3d': {
        // under an operator or alone Excel answers #REF! for it
        if (!rangeOk) return b.value(E.REF)
        const x = rect(n)
        const [i, j] = [ctx.sheet(n.first), ctx.sheet(n.last)]
        const list: Rng[] = []
        for (let k = Math.min(i, j); k <= Math.max(i, j); k++) list.push(new Rng(k, x.r1, x.c1, x.r2, x.c2))
        info.refs.push(...list)
        return b.areas(list)
      }
      case 'name': {
        const d = resolve(n, scope)
        if (!d) return b.value(E.NAME)
        // a name that uses itself, through other names or not
        if (inside.has(d)) return b.value(E.CYCLE)
        const outer: [T[], number] = [uses, scope]
        ;[uses, scope] = [[], -1]
        inside.add(d)
        const code = go(d.ast, rangeOk)
        inside.delete(d)
        ;[uses, scope] = outer
        return code
      }
      case 'at':
        return value(b.at(go(n.a, true), r, c), n, rangeOk)
      case 'union':
        return value(
          b.union(n.items.map((x) => go(x, true))),
          n,
          rangeOk,
        )
      case 'paren':
        return go(n.a, rangeOk)
      case 'un':
        // unary plus keeps a reference (SUM(+A1:A3) sums the range)
        return n.op === '+' ? go(n.a) : b.unary('neg', go(n.a), many(n.a))
      case 'pct':
        return b.unary('pct', go(n.a), many(n.a))
      case 'bin': {
        const x = go(n.a)
        return b.binary(n.op, x, go(n.b), many(n.a) || many(n.b))
      }
      case 'fn': {
        if (ctx.area && !n.args.length && (n.name === 'ROW' || n.name === 'COLUMN')) return b.array(lines(ctx.area, n.name === 'ROW'))
        const spec = FUNCTIONS[n.name]
        if (!spec) {
          info.external = true
          return b.external(canonical(n))
        }
        if (spec.volatile) info.volatile = true
        if (n.name === 'SUBTOTAL' || n.name === 'AGGREGATE') info.nested = true
        // a reference across sheets is passed to any argument, which takes it or answers its error
        const args = n.args.map((a, i) => (a.k === 'missing' ? omitted(spec.omitted(i)) : go(a, spec.range(i) || a.k === 'ref3d')))
        let code: T
        if (spec.lazy) {
          // lifted when a lifting argument is an array
          code = b.lazy(n.name, args, n.args.flatMap((a, i) => (many(a) && spec.liftAt(i, n.args.length) ? [i] : [])))
        } else {
          // RAND gets its cell
          if (spec.cell) args.unshift(b.value(s), b.value(r), b.value(c))
          code = b.call(n.name, args, n.args.some((a, i) => many(a) && spec.mode(i) !== 'whole'))
        }
        // a reference answered where one value goes is read here; a range position takes it as a reference
        return spec.refs ? value(code, n, rangeOk) : code
      }
      case 'missing':
        return b.value(null)
      case 'let':
        // the values are kept as references; the result is read where one value goes
        return b.let(n.binds.length, (local) => {
          n.binds.forEach((x, i) => (uses[x.slot] = local(i)))
          const values = n.binds.map((x) => go(x.value, true))
          return { values, body: n.body ? go(n.body, rangeOk) : value(uses[n.binds[n.binds.length - 1].slot], n, rangeOk) }
        })
      case 'local':
        return value(uses[n.slot], n, rangeOk)
    }
  }
  const code = go(ast)
  return { ...info, code }
}
