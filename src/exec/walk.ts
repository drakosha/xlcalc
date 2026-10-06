/**
 * The one walk over a formula's AST that both execution modes share. References, names, shapes
 * (`arrays/shape.ts`), function specs and the formula's cell are resolved here, together with the formula's
 * references and flags; a backend (`js.ts` for `eval`, `closures.ts`) only builds the code of each node
 * from what the walk decided, so the two modes cannot part on a decision.
 */
import { shapes } from '../arrays/shape.js'
import { Arr } from '../arrays/value.js'
import { FUNCTIONS } from '../functions.js'
import { canonical, type Ast, type BinOp } from '../syntax.js'
import { E, Rng, type V } from '../values.js'

/** What the walk needs to know of the book, seen from the formula's sheet. */
export interface CodegenContext {
  /** The index of a sheet (null: the formula's own). */
  sheet(name: string | null): number
  /** A name visible from the formula's sheet: its reference (or `#REF!` when lost), and whether it has several cells. */
  name(name: string): { ast: Extract<Ast, { k: 'ref' | 'err' }>; area: boolean } | undefined
}

export type Operator = 'add' | 'sub' | 'mul' | 'div' | 'pow' | 'cat'

/** The runtime's arithmetic and text operators by symbol; comparisons go through `cmp` (lifted: `C`). */
export const OPERATORS: Partial<Record<BinOp, Operator>> = { '+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat' }

/**
 * Code of one node in a mode. `lift`: an operand or a one-value argument can be an array, so the node goes
 * through the lifting operators (`L1`, `L2`) or callers (`FL`, `LZ`); otherwise it is plain scalar code.
 */
export interface Backend<T> {
  /** A literal: a number, text, a boolean or an error. */
  value(v: V): T
  /** An array constant, built once per compiled formula. */
  array(a: Arr): T
  cell(s: number, r: number, c: number): T
  /** A constant range, built once per compiled formula; `one`: a single cell in a one-value position. */
  range(g: Rng, one: boolean): T
  unary(op: 'neg' | 'pct', a: T, lift: boolean): T
  binary(op: BinOp, a: T, b: T, lift: boolean): T
  call(name: string, args: T[], lift: boolean): T
  /** A lazy function (IF, IFERROR): each argument computed on demand; `check`, the positions that can lift it. */
  lazy(name: string, args: T[], check: number[]): T
  /** The value of a call that may answer a one-cell reference (CHOOSE, IF) where one value goes. */
  read(code: T): T
  /** An external function, by the canonical text of its call. */
  external(text: string): T
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

/** Walks a formula at sheet `s`, row `r`, column `c`: its information and the code of its root in `b`. */
export function walk<T>(ctx: CodegenContext, ast: Ast, s: number, r: number, c: number, b: Backend<T>): FormulaInfo & { code: T } {
  const info: FormulaInfo = { refs: [], volatile: false, nested: false, external: false }
  const deref = (n: Ast): Ast => (n.k === 'name' ? (ctx.name(n.name)?.ast ?? n) : n)
  const shape = shapes(ast, (name) => ctx.name(name)?.area)
  const many = (n: Ast) => (shape.get(n) ?? 0) > 0
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
        const g = new Rng(ctx.sheet(n.sheet), n.ref.r1, n.ref.c1, n.ref.r2, n.ref.c2)
        info.refs.push(g)
        if (n.ref.kind === 'cell' && !rangeOk) return b.cell(g.s, g.r1, g.c1)
        return b.range(g, !rangeOk && g.r2 === g.r1 && g.c2 === g.c1)
      }
      case 'name': {
        const d = deref(n)
        return d.k === 'name' ? b.value(E.NAME) : go(d, rangeOk)
      }
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
        const spec = FUNCTIONS[n.name]
        if (!spec) {
          info.external = true
          return b.external(canonical(n))
        }
        if (spec.volatile) info.volatile = true
        if (n.name === 'SUBTOTAL' || n.name === 'AGGREGATE') info.nested = true
        const args = n.args.map((a, i) => go(a, spec.range(i)))
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
        return spec.refs && !rangeOk && !many(n) ? b.read(code) : code
      }
    }
  }
  const code = go(ast)
  return { ...info, code }
}
